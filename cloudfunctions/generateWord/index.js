const cloud = require('wx-server-sdk');
const {
	Document,
	Packer,
	Paragraph,
	TextRun,
	HeadingLevel,
	AlignmentType,
	BorderStyle,
	Table,
	TableRow,
	TableCell,
	WidthType,
	PageBreak,
	Header,
	ImageRun,
} = require('docx');
const { Math: OfficeMath, MathRun, latexToMath } = require('docx/math');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });

// 图片公式仅用于手机端：微信内置的 DOCX 阅读器不会稳定地绘制 Office Math
// (OMML)，但会绘制普通 PNG。渲染模块延迟加载，避免它的运行环境问题影响
// 电脑端的可编辑公式导出。
let formulaRenderer;

function latexForImage(latex) {
	return String(latex || '')
		.replace(/\r\n?/g, '\n')
		.trim()
		// 从 Markdown/聊天窗口复制时，反斜杠常被额外转义一次。
		.replace(/\\\\(?=[A-Za-z])/g, '\\');
}

async function getFormulaRenderer() {
	if (formulaRenderer) return formulaRenderer;
	try {
		const fs = require('fs');
		const { mathjax } = require('mathjax-full/js/mathjax.js');
		const { TeX } = require('mathjax-full/js/input/tex.js');
		const { SVG } = require('mathjax-full/js/output/svg.js');
		const { liteAdaptor } = require('mathjax-full/js/adaptors/liteAdaptor.js');
		const { RegisterHTMLHandler } = require('mathjax-full/js/handlers/html.js');
		const { AllPackages } = require('mathjax-full/js/input/tex/AllPackages.js');
		const { Resvg, initWasm } = require('@resvg/resvg-wasm');

		// resvg 的 WASM 版没有 sharp 之类的本机二进制依赖，适合云函数运行时。
		await initWasm(fs.readFileSync(require.resolve('@resvg/resvg-wasm/index_bg.wasm')));
		const adaptor = liteAdaptor();
		RegisterHTMLHandler(adaptor);
		const document = mathjax.document('', {
			InputJax: new TeX({ packages: AllPackages }),
			OutputJax: new SVG({ fontCache: 'none' }),
		});

		formulaRenderer = (latex) => {
			const node = document.convert(latexForImage(latex), {
				display: true,
				em: 16,
				ex: 8,
				containerWidth: 80 * 16,
			});
			// MathJax 返回的是 <mjx-container><svg>...</svg></mjx-container>。
			// resvg 只接受以 <svg> 为根的 XML，不能把外层容器传进去。
			const wrappedSvg = adaptor.outerHTML(node);
			const svgMatch = wrappedSvg.match(/<svg\b[\s\S]*<\/svg>/i);
			if (!svgMatch) throw new Error('公式排版结果中没有 SVG 图像');
			const svg = svgMatch[0];
			const rendered = new Resvg(svg, {
				background: 'rgba(255,255,255,0)',
				fitTo: { mode: 'zoom', value: 2 },
			}).render();
			return { data: rendered.asPng(), width: rendered.width, height: rendered.height };
		};
		return formulaRenderer;
	} catch (error) {
		throw new Error(`手机公式图片渲染器未就绪：${error.message || error}`);
	}
}

async function buildFormulaImages(blocks) {
	const renderer = await getFormulaRenderer();
	const images = new Map();
	const add = (latex) => {
		const source = latexForImage(latex);
		if (source) images.set(source, null);
	};
	const visitText = (text) => {
		const source = decorateBareInlineMath(String(text || ''));
		const pattern = /\\\(([\s\S]+?)\\\)|\\\[([\s\S]+?)\\\]|\$\$([\s\S]+?)\$\$|\$([^$\n]+?)\$/g;
		let match;
		while ((match = pattern.exec(source))) add(match[1] || match[2] || match[3] || match[4]);
	};
	for (const block of blocks) {
		if (block.type === 'formula') add(block.content);
		else if (block.type === 'table' && Array.isArray(block.content)) block.content.forEach((row) => row.forEach(visitText));
		else if (block.type === 'list' && Array.isArray(block.content)) block.content.forEach(visitText);
		else visitText(block.content);
	}
	for (const latex of images.keys()) images.set(latex, renderer(latex));
	return images;
}


const headingLevels = [
	HeadingLevel.HEADING_1,
	HeadingLevel.HEADING_2,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
];

// 模板必须影响实际 DOCX，而不只是选择页上的名称。
const TEMPLATE_LAYOUTS = {
	homework: {
		name: '课程作业', body: { font: '宋体', size: 24 }, heading: { font: '黑体', size: 32, bold: true },
		bodySpacing: 360, headingSpacing: 280, numberHeadings: false, header: '',
	},
	'lab-report': {
		name: '实验报告', body: { font: '宋体', size: 24 }, heading: { font: '黑体', size: 32, bold: true },
		bodySpacing: 420, headingSpacing: 320, numberHeadings: true, header: '实验报告',
	},
	'study-note': {
		name: '学术笔记', body: { font: '仿宋_GB2312', size: 22 }, heading: { font: '微软雅黑', size: 30, bold: true, color: '285943' },
		bodySpacing: 360, headingSpacing: 240, numberHeadings: false, header: '学术笔记',
	},
};

function layoutFor(templateId) {
	return TEMPLATE_LAYOUTS[templateId] || TEMPLATE_LAYOUTS.homework;
}

/**
 * 作为 LaTeX 解析失败时的兜底文本。正常路径必须写入 OMML，才能在 Word/WPS
 * 中以公式对象打开并继续编辑。
 */
function formulaText(latex, tagMode) {
	// 保留 array/aligned 中的真实换行：手机端虽不能渲染 OMML，但仍应按行展示。
	let value = String(latex || '').replace(/\r\n?/g, '\n').trim();
	// 部分复制来源会把 \begin、\end、\tag 写成双反斜杠；这不是数组的换行，
	// 先还原命令本身，避免手机预览泄露原始控制标签。
	value = value.replace(/\\\\(?=(?:begin|end|tag)\b)/g, '\\');
	// 必须在 \le 等短命令转换前移除尺寸控制命令，否则 \left 会被误识别为 \le + ft。
	value = value.replace(/\\left|\\right|\\bigg|\\Bigg|\\big(?!cup|cap)|\\Big/g, '');
	// 展开常见的嵌套分式。有限循环可避免不受控的正则替换。
	for (let i = 0; i < 8; i += 1) {
		const next = value
			// 分子/分母常带 H_{in} 这类一层嵌套花括号，不能只匹配纯文本分式。
			.replace(/\\(?:dfrac|tfrac|cfrac|frac)\{((?:[^{}]|\{[^{}]*\})*)\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, '($1)/($2)')
			.replace(/\\sqrt\{([^{}]*)\}/g, '√($1)')
			.replace(/\\mathcal\{L\}/g, 'ℒ')
			.replace(/\\(?:text|mathrm|mathbf|mathit|mathcal|mathbb|boldsymbol|operatorname)\{([^{}]*)\}/g, '$1');
		if (next === value) break;
		value = next;
	}

	const commands = [
		['\\Rightarrow', '⇒'], ['\\rightarrow', '→'], ['\\leftarrow', '←'],
		['\\langle', '⟨'], ['\\rangle', '⟩'], ['\\lceil', '⌈'], ['\\rceil', '⌉'],
		['\\times', '×'], ['\\cdot', '·'], ['\\leq', '≤'], ['\\geq', '≥'],
		['\\neq', '≠'], ['\\ne', '≠'], ['\\le', '≤'], ['\\ge', '≥'],
		['\\pm', '±'], ['\\infty', '∞'], ['\\bigcup', '⋃'], ['\\bigcap', '⋂'], ['\\sum', '∑'], ['\\prod', '∏'], ['\\int', '∫'],
		['\\forall', '∀'], ['\\exists', '∃'], ['\\emptyset', '∅'], ['\\notin', '∉'], ['\\in', '∈'], ['\\cup', '∪'], ['\\cap', '∩'],
		['\\cdots', '⋯'], ['\\ldots', '…'], ['\\dots', '…'],
		['\\alpha', 'α'], ['\\beta', 'β'], ['\\gamma', 'γ'], ['\\delta', 'δ'], ['\\epsilon', 'ε'], ['\\varepsilon', 'ε'], ['\\zeta', 'ζ'], ['\\eta', 'η'], ['\\theta', 'θ'], ['\\vartheta', 'ϑ'],
		['\\iota', 'ι'], ['\\kappa', 'κ'], ['\\lambda', 'λ'], ['\\mu', 'μ'], ['\\nu', 'ν'], ['\\xi', 'ξ'], ['\\pi', 'π'], ['\\rho', 'ρ'], ['\\sigma', 'σ'], ['\\tau', 'τ'], ['\\upsilon', 'υ'], ['\\phi', 'φ'], ['\\varphi', 'ϕ'], ['\\chi', 'χ'], ['\\psi', 'ψ'], ['\\omega', 'ω'],
		['\\odot', '⊙'], ['\\partial', '∂'], ['\\nabla', '∇'], ['\\approx', '≈'], ['\\equiv', '≡'],
	];
	commands.forEach(([from, to]) => { value = value.split(from).join(to); });
	value = value
		// 手机兼容模式必须去掉所有控制命令，避免微信预览把 LaTeX 原文直接显示出来。
		// array/aligned 的 `&` 是列对齐标记、`\\` 是行结束标记，转成普通 Word 换行。
		.replace(/\\begin\{(?:array|aligned|align\*?|cases|matrix|pmatrix|bmatrix|vmatrix|Vmatrix)\}(?:\{[^{}]*\})?/g, '')
		.replace(/\\end\{(?:array|aligned|align\*?|cases|matrix|pmatrix|bmatrix|vmatrix|Vmatrix)\}/g, '')
		.replace(/\\\\(?:\[[^\]]*\])?/g, '\n')
		.replace(/&/g, ' ')
		.replace(/\\(?:under|over)line\{((?:[^{}]|\{[^{}]*\})*)\}/g, '$1')
		// 原始公式自带 \\tag 时，在未启用“显示公式编号”时仍应保留编号；
		// 启用该选项则由右侧编号列统一绘制，避免出现两个编号。
		// 同时兼容部分识别结果省略花括号而得到的 `\\tag1`。
		.replace(/\\tag\s*(?:\{([^{}]*)\}|(\d+))/g, (_, bracedTag, bareTag) => {
			const tag = bracedTag || bareTag;
			return tagMode === 'hide' ? '' : ` (${tag})`;
		})
		.replace(/\\quad|\\qquad/g, ' ')
		.replace(/\\(?:min|max|cos|sin|tan|cot|log|ln|exp)\b/g, (match) => match.slice(1))
		.replace(/\\Gamma/g, 'Γ').replace(/\\Delta/g, 'Δ').replace(/\\Theta/g, 'Θ').replace(/\\Lambda/g, 'Λ').replace(/\\Xi/g, 'Ξ').replace(/\\Pi/g, 'Π').replace(/\\Sigma/g, 'Σ').replace(/\\Phi/g, 'Φ').replace(/\\Psi/g, 'Ψ').replace(/\\Omega/g, 'Ω')
		.replace(/\\(?:widehat|hat)\s*\{?\s*([A-Za-z])\s*\}?/g, '$1̂')
		.replace(/\\[,!;:\s]+/g, ' ')
		.replace(/\\([{}])/g, '$1');

	// H_{out} / x^2 转为 Unicode 上下标；多字符上下标不会因删除花括号而丢失。
	value = value.replace(/_\{([^{}]+)\}/g, (_, subscript) => `₍${subscript}₎`);
	value = value.replace(/\^\{([^{}]+)\}/g, (_, superscript) => `^(${superscript})`);
	value = value.replace(/_([A-Za-z0-9])/g, (_, subscript) => `₍${subscript}₎`);
	value = value.replace(/\^([A-Za-z0-9])/g, (_, superscript) => `^(${superscript})`);
	return value.replace(/[{}]/g, '').split('\n').map((line) => line.replace(/[ \t]+/g, ' ').trim()).filter(Boolean).join('\n') || '[空公式]';
}

function editableMath(latex) {
	const fallback = new TextRun({
		text: formulaText(latex),
		font: 'Cambria Math',
		// 微信文档预览器会跳过 m:oMath。它通常不会处理“隐藏文字”标记，因而
		// 仍可显示此兼容文本；Word/WPS 则只显示前面的可编辑公式对象。
		vanish: true,
	});
	try {
		return [new OfficeMath({ children: latexToMath(String(latex || '').trim()) }), fallback];
	} catch (error) {
		// 单个不受支持的命令不能让整份文档导出失败；此时仍保留可读内容。
		return [new OfficeMath({ children: [new MathRun(formulaText(latex))] }), fallback];
	}
}

function formulaImageRun(latex, formulaImages) {
	const image = formulaImages && formulaImages.get(latexForImage(latex));
	if (!image) throw new Error('公式图片缺失，请重新生成文档');
	// Word 的尺寸单位是 CSS 像素。限制宽度可避免长公式撑破页面，同时保持比例。
	const width = Math.max(24, Math.min(520, Math.round(image.width / 2)));
	const height = Math.max(18, Math.min(720, Math.round((image.height / image.width) * width)));
	return new ImageRun({ data: image.data, type: 'png', transformation: { width, height } });
}

function mathRuns(latex, formulaMode, tagMode, formulaImages) {
	if (formulaMode === 'image') return [formulaImageRun(latex, formulaImages)];
	// 微信手机端预览器不渲染 OMML；改用常规文字 Run 才能保证可见。
	if (formulaMode === 'compatible') {
		const lines = formulaText(latex, tagMode).split('\n');
		return lines.map((line, index) => new TextRun({ text: line, font: 'Cambria Math', break: index || undefined }));
	}
	return editableMath(latex);
}

function normalTextRun(text, options) {
	// 未被识别为完整公式的正文，也至少转换不依赖分组的常见符号，避免出现 \times。
	const visible = String(text || '')
		.replace(/\\times/g, '×').replace(/\\cdot/g, '·')
		.replace(/\\Rightarrow/g, '⇒').replace(/\\rightarrow/g, '→').replace(/\\leq|\\le/g, '≤').replace(/\\geq|\\ge/g, '≥')
		.replace(/\\,/g, ' ').replace(/\\;/g, ' ');
	return new TextRun(Object.assign({ text: visible }, options || {}));
}

function decorateBareInlineMath(source) {
	let value = String(source || '');
	// 先保护已经带定界符的公式，避免 H_{out} 被二次包裹成 \(\(H...\)\)。
	const protectedTokens = [];
	value = value.replace(/\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$\$[\s\S]+?\$\$|\$[^$\n]+?\$/g, (match) => {
		const key = `@@MATH_TOKEN_${protectedTokens.length}@@`;
		protectedTokens.push(match);
		return key;
	});
	// 将大模型常省略分隔符的 H_{out}=H_{in} 包为公式 token。
	value = value.replace(/\b[A-Za-z]_\{[^{}]+\}(?:\s*=\s*[A-Za-z0-9_{}^\\.,:+\-*/=<>()[\]\s]+)?/g, (match) => `\\(${match.trim()}\\)`);
	// 处理独立出现的 \boldsymbol{...}、\frac{...}{...} 等片段。
	value = value.replace(/\\(?:boldsymbol|mathbf|mathit|mathrm)\{[^{}]+\}/g, (match) => `\\(${match}\\)`);
	value = value.replace(/\\(?:d?frac|tfrac|cfrac)\{(?:[^{}]|\{[^{}]*\})*\}\{(?:[^{}]|\{[^{}]*\})*\}/g, (match) => `\\(${match}\\)`);
	return value.replace(/@@MATH_TOKEN_(\d+)@@/g, (_, index) => protectedTokens[Number(index)]);
}

function inlineRuns(text, runOptions, formulaMode, formulaImages) {
	const source = decorateBareInlineMath(String(text || '').replace(/\r\n?/g, ' '));
	const runs = [];
	const pattern = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)/g;
	let lastIndex = 0;
	let match = pattern.exec(source);
	while (match) {
		if (match.index > lastIndex) {
			runs.push(normalTextRun(source.slice(lastIndex, match.index), runOptions));
		}

		const token = match[0];
		if (token.startsWith('**')) {
			runs.push(normalTextRun(token.slice(2, -2), Object.assign({}, runOptions, { bold: true })));
		} else if (token.startsWith('*')) {
			runs.push(normalTextRun(token.slice(1, -1), Object.assign({}, runOptions, { italics: true })));
		} else if (token.startsWith('`')) {
			runs.push(normalTextRun(token.slice(1, -1), Object.assign({}, runOptions, { font: 'Courier New' })));
		} else if (token.startsWith('$$')) {
			runs.push(...mathRuns(token.slice(2, -2), formulaMode, undefined, formulaImages));
		} else if (token.startsWith('\\(')) {
			runs.push(...mathRuns(token.slice(2, -2), formulaMode, undefined, formulaImages));
		} else if (token.startsWith('\\[')) {
			runs.push(...mathRuns(token.slice(2, -2), formulaMode, undefined, formulaImages));
		} else if (token.startsWith('$')) {
			runs.push(...mathRuns(token.slice(1, -1), formulaMode, undefined, formulaImages));
		}

		lastIndex = pattern.lastIndex;
		match = pattern.exec(source);
	}

	if (lastIndex < source.length) {
		runs.push(normalTextRun(source.slice(lastIndex), runOptions));
	}

	return runs.length ? runs : [new TextRun({ text: '' })];
}

function paragraphForBlock(block, layout, headingPrefix, formulaMode, formulaImages) {
	if (block.type === 'heading') {
		return new Paragraph({
			children: inlineRuns(`${headingPrefix || ''}${block.content}`, layout.heading, formulaMode, formulaImages),
			heading: headingLevels[Math.max(0, Math.min((block.level || 1) - 1, 2))],
			spacing: { before: layout.headingSpacing, after: layout.headingSpacing },
		});
	}

	if (block.type === 'paragraph') {
		return new Paragraph({ children: inlineRuns(block.content, layout.body, formulaMode, formulaImages), spacing: { after: layout.bodySpacing, line: layout.bodySpacing } });
	}

	if (block.type === 'quote') {
		return new Paragraph({
			children: inlineRuns(block.content, layout.body, formulaMode, formulaImages),
			indent: { left: 720 },
			border: {
				left: {
					color: 'cccccc',
					space: 1,
					style: BorderStyle.SINGLE,
					size: 12,
				},
			},
			spacing: { after: layout.bodySpacing, line: layout.bodySpacing },
		});
	}

	if (block.type === 'code') {
		const lines = String(block.content || '').split('\n');
		return lines.flatMap((line) => [
			new Paragraph({
				children: [new TextRun({ text: line || ' ', font: 'Courier New', size: layout.body.size, color: '303030' })],
				shading: { fill: 'F2F3F0' },
				spacing: { after: 100, line: 280 },
			}),
		]);
	}

	if (block.type === 'formula') return formulaParagraphs(block.content, 0, layout, formulaMode, formulaImages);

	if (block.type === 'list') {
		return (block.content || []).map((item, index) => {
			const runs = inlineRuns(item, layout.body, formulaMode, formulaImages);
			if (block.ordered) {
				runs.unshift(new TextRun({ text: `${index + 1}. ` }));
			}
			return new Paragraph({
				children: runs,
				bullet: block.ordered ? undefined : { level: 0 },
				spacing: { after: 80, line: layout.bodySpacing },
			});
		});
	}

	return new Paragraph({ children: inlineRuns(block.content || '', layout.body, formulaMode, formulaImages), spacing: { after: layout.bodySpacing, line: layout.bodySpacing } });
}

function formulaParagraphs(latex, number, layout, formulaMode, formulaImages) {
	const formula = new Paragraph({
		children: mathRuns(latex, formulaMode, number ? 'hide' : 'show', formulaImages),
		alignment: AlignmentType.CENTER,
		spacing: { after: layout.bodySpacing },
	});
	if (!number) return formula;

	// 三列布局让编号靠右，同时不把公式从页面中央挤开。
	return new Table({
		width: { size: 100, widthType: WidthType.PERCENTAGE },
		borders: { top: { style: BorderStyle.NIL }, bottom: { style: BorderStyle.NIL }, left: { style: BorderStyle.NIL }, right: { style: BorderStyle.NIL }, insideHorizontal: { style: BorderStyle.NIL }, insideVertical: { style: BorderStyle.NIL } },
		rows: [new TableRow({ children: [
			new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, children: [new Paragraph('')] }),
			new TableCell({ width: { size: 70, type: WidthType.PERCENTAGE }, children: [formula] }),
			new TableCell({ width: { size: 15, type: WidthType.PERCENTAGE }, children: [new Paragraph({ text: `(${number})`, alignment: AlignmentType.CENTER })] }),
		] })],
	});
}

function nextHeadingPrefix(block, layout, levels) {
	if (!layout.numberHeadings || block.type !== 'heading') return '';
	const level = Math.max(1, Math.min(block.level || 1, 3));
	levels[level - 1] = (levels[level - 1] || 0) + 1;
	for (let index = level; index < levels.length; index += 1) levels[index] = 0;
	return `${levels.slice(0, level).map((value) => value || 0).join('.')}. `;
}

function sectionFor(children, layout) {
	const section = { properties: {}, children };
	if (layout.header) {
		section.headers = {
			default: new Header({
				children: [new Paragraph({
					children: [new TextRun({ text: layout.header, font: layout.body.font, size: 20, color: '6E6E6A' })],
					alignment: AlignmentType.RIGHT,
				})],
			}),
		};
	}
	return section;
}

async function generateWordFromBlocks(blocks) {
	const layout = TEMPLATE_LAYOUTS.homework;
	const children = [];
	for (const block of blocks) {
		const paragraph = paragraphForBlock(block, layout, '', 'compatible');
		children.push(...(Array.isArray(paragraph) ? paragraph : [paragraph]));
	}

	const doc = new Document({
		sections: [
			{
				properties: {},
				children,
			},
		],
	});

	return Packer.toBuffer(doc);
}

exports.main = async (event) => {
	const d = event.draft || {};
	const blocks = Array.isArray(d.blocks) ? d.blocks : [];
	if (!d.id || blocks.length > 1000) throw new Error('无效的导出内容');

	const children = [];
	const exportOptions = d.exportOptions || event.exportOptions || {};
	const layout = layoutFor(d.templateId);
	// 云函数也以兼容格式为默认值，避免旧客户端或第三方调用在微信预览中
	// 生成不可见的 OMML 公式。只有显式请求 editable 时才生成可编辑公式对象。
	const formulaMode = ['editable', 'image'].includes(event.formulaMode) ? event.formulaMode : 'compatible';
	// 仅手机图片模式初始化渲染器；桌面可编辑公式完全不受图片依赖影响。
	const formulaImages = formulaMode === 'image' ? await buildFormulaImages(blocks) : null;
	if (exportOptions.cover) {
		children.push(
			new Paragraph({ children: [new TextRun({ text: d.title || '未命名文档', font: layout.heading.font, size: 44, bold: true })], alignment: AlignmentType.CENTER, spacing: { before: 1800, after: 360 } }),
			new Paragraph({ children: [new TextRun({ text: layout.name, font: layout.body.font, size: 28 })], alignment: AlignmentType.CENTER }),
			new Paragraph({ children: [new PageBreak()] })
		);
	}

	let formulaNumber = 0;
	const headingLevels = [0, 0, 0];
	for (const block of blocks) {
		if (block.type === 'table' && Array.isArray(block.content)) {
			children.push(
				new Table({
					width: { size: 100, widthType: WidthType.PERCENTAGE },
					rows: block.content.map((row, rowIndex) => new TableRow({ children: row.map((cell) => new TableCell({ shading: rowIndex === 0 ? { fill: 'EDF3EE' } : undefined, children: [new Paragraph({ children: inlineRuns(String(cell), Object.assign({}, layout.body, rowIndex === 0 ? { bold: true } : {}), formulaMode, formulaImages), spacing: { line: layout.bodySpacing } })] })) })),
				})
			);
		} else if (block.type === 'formula') {
			formulaNumber += 1;
			children.push(formulaParagraphs(block.content, exportOptions.formulaNumbers ? formulaNumber : 0, layout, formulaMode, formulaImages));
		} else {
			const paragraph = paragraphForBlock(block, layout, nextHeadingPrefix(block, layout, headingLevels), formulaMode, formulaImages);
			children.push(...(Array.isArray(paragraph) ? paragraph : [paragraph]));
		}
	}

	const buffer = await Packer.toBuffer(new Document({ sections: [sectionFor(children, layout)] }));
	const openid = cloud.getWXContext().OPENID;
	const safeTitle = String(d.title || '未命名文档').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) || '未命名文档';
	const fileName = `${safeTitle}.docx`;
	// 云端路径仍带时间戳以避免同名文档覆盖；只把友好的文件名返回给前端。
	const cloudPath = `${openid}/${d.id}/${Date.now()}_${fileName}`;
	const upload = await cloud.uploadFile({ cloudPath, fileContent: buffer });
	return { success: true, fileID: upload.fileID, fileName, generatorVersion: 5, formulaMode };
};
