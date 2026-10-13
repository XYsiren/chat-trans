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
} = require('docx');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });

const headingLevels = [
	HeadingLevel.HEADING_1,
	HeadingLevel.HEADING_2,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
];

/**
 * 微信里打开 docx 使用的是内置预览器，它不会稳定渲染 OOXML 的 m:oMath。
 * 以前直接写入原生公式对象时，Word 中可编辑、但在小程序里会显示为空白。
 * 因此导出链路优先生成跨端可见的数学文本；content 仍保留原始 LaTeX，方便
 * 后续增加“桌面 Word 原生公式”格式时复用。
 */
function formulaText(latex) {
	let value = String(latex || '').replace(/\r\n?/g, ' ').trim();
	// 展开常见的嵌套分式。有限循环可避免不受控的正则替换。
	for (let i = 0; i < 8; i += 1) {
		const next = value
			// 分子/分母常带 H_{in} 这类一层嵌套花括号，不能只匹配纯文本分式。
			.replace(/\\(?:dfrac|tfrac|cfrac|frac)\{((?:[^{}]|\{[^{}]*\})*)\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, '($1)/($2)')
			.replace(/\\sqrt\{([^{}]*)\}/g, '√($1)')
			.replace(/\\(?:text|mathrm|mathbf|mathit|boldsymbol|operatorname)\{([^{}]*)\}/g, '$1');
		if (next === value) break;
		value = next;
	}

	const commands = [
		['\\Rightarrow', '⇒'], ['\\rightarrow', '→'], ['\\leftarrow', '←'],
		['\\times', '×'], ['\\cdot', '·'], ['\\leq', '≤'], ['\\geq', '≥'],
		['\\neq', '≠'], ['\\ne', '≠'], ['\\le', '≤'], ['\\ge', '≥'],
		['\\pm', '±'], ['\\infty', '∞'], ['\\sum', '∑'], ['\\prod', '∏'], ['\\int', '∫'],
		['\\forall', '∀'], ['\\exists', '∃'], ['\\in', '∈'], ['\\cup', '∪'], ['\\cap', '∩'],
		['\\alpha', 'α'], ['\\beta', 'β'], ['\\gamma', 'γ'], ['\\delta', 'δ'], ['\\theta', 'θ'],
	];
	commands.forEach(([from, to]) => { value = value.split(from).join(to); });
	value = value
		.replace(/\\left|\\right|\\bigg|\\Bigg|\\big|\\Big/g, '')
		.replace(/\\[,!;:\s]+/g, ' ')
		.replace(/\\([{}])/g, '$1');

	// H_{out} / x^2 转为 Unicode 上下标；多字符上下标不会因删除花括号而丢失。
	value = value.replace(/_\{([^{}]+)\}/g, (_, subscript) => `₍${subscript}₎`);
	value = value.replace(/\^\{([^{}]+)\}/g, (_, superscript) => `^(${superscript})`);
	return value.replace(/[{}]/g, '').replace(/\s+/g, ' ').trim() || '[空公式]';
}

function mathRun(latex) {
	return new TextRun({ text: formulaText(latex), font: 'Cambria Math' });
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

function inlineRuns(text) {
	const source = decorateBareInlineMath(String(text || '').replace(/\r\n?/g, ' '));
	const runs = [];
	const pattern = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)/g;
	let lastIndex = 0;
	let match = pattern.exec(source);
	while (match) {
		if (match.index > lastIndex) {
			runs.push(normalTextRun(source.slice(lastIndex, match.index)));
		}

		const token = match[0];
		if (token.startsWith('**')) {
			runs.push(normalTextRun(token.slice(2, -2), { bold: true }));
		} else if (token.startsWith('*')) {
			runs.push(normalTextRun(token.slice(1, -1), { italics: true }));
		} else if (token.startsWith('`')) {
			runs.push(normalTextRun(token.slice(1, -1), { font: 'Courier New' }));
		} else if (token.startsWith('$$')) {
			runs.push(mathRun(token.slice(2, -2)));
		} else if (token.startsWith('\\(')) {
			runs.push(mathRun(token.slice(2, -2)));
		} else if (token.startsWith('\\[')) {
			runs.push(mathRun(token.slice(2, -2)));
		} else if (token.startsWith('$')) {
			runs.push(mathRun(token.slice(1, -1)));
		}

		lastIndex = pattern.lastIndex;
		match = pattern.exec(source);
	}

	if (lastIndex < source.length) {
		runs.push(normalTextRun(source.slice(lastIndex)));
	}

	return runs.length ? runs : [new TextRun({ text: '' })];
}

function paragraphForBlock(block) {
	if (block.type === 'heading') {
		return new Paragraph({
			children: inlineRuns(block.content),
			heading: headingLevels[Math.max(0, Math.min((block.level || 1) - 1, 2))],
			spacing: { after: 200 },
		});
	}

	if (block.type === 'paragraph') {
		return new Paragraph({ children: inlineRuns(block.content), spacing: { after: 200 } });
	}

	if (block.type === 'quote') {
		return new Paragraph({
			children: inlineRuns(block.content),
			indent: { left: 720 },
			border: {
				left: {
					color: 'cccccc',
					space: 1,
					style: BorderStyle.SINGLE,
					size: 12,
				},
			},
			spacing: { after: 200 },
		});
	}

	if (block.type === 'code') {
		const lines = String(block.content || '').split('\n');
		return lines.flatMap((line) => [
			new Paragraph({
				children: [new TextRun({ text: line || ' ', font: 'Courier New' })],
				spacing: { after: 100 },
			}),
		]);
	}

	if (block.type === 'formula') {
		return new Paragraph({
			children: [mathRun(block.content)],
			alignment: AlignmentType.CENTER,
			spacing: { after: 200 },
		});
	}

	if (block.type === 'list') {
		return (block.content || []).map((item, index) => {
			const runs = inlineRuns(item);
			if (block.ordered) {
				runs.unshift(new TextRun({ text: `${index + 1}. ` }));
			}
			return new Paragraph({
				children: runs,
				bullet: block.ordered ? undefined : { level: 0 },
				spacing: { after: 80 },
			});
		});
	}

	return new Paragraph({ children: inlineRuns(block.content || ''), spacing: { after: 200 } });
}

async function generateWordFromBlocks(blocks) {
	const children = [];
	for (const block of blocks) {
		const paragraph = paragraphForBlock(block);
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
	if (d.exportOptions && d.exportOptions.cover) {
		children.push(
			new Paragraph({ text: d.title || '未命名文档', heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }),
			new Paragraph({ text: '' })
		);
	}

	for (const block of blocks) {
		if (block.type === 'table' && Array.isArray(block.content)) {
			children.push(
				new Table({
					width: { size: 100, widthType: WidthType.PERCENTAGE },
					rows: block.content.map((row) => new TableRow({ children: row.map((cell) => new TableCell({ children: [new Paragraph(String(cell))] })) })),
				})
			);
		} else {
			const paragraph = paragraphForBlock(block);
			children.push(...(Array.isArray(paragraph) ? paragraph : [paragraph]));
		}
	}

	const buffer = await Packer.toBuffer(new Document({ sections: [{ children }] }));
	const openid = cloud.getWXContext().OPENID;
	const safeTitle = String(d.title || '未命名文档').trim().replace(/[\\/:*?"<>|]/g, '_').slice(0, 80) || '未命名文档';
	const fileName = `${safeTitle}.docx`;
	// 云端路径仍带时间戳以避免同名文档覆盖；只把友好的文件名返回给前端。
	const cloudPath = `${openid}/${d.id}/${Date.now()}_${fileName}`;
	const upload = await cloud.uploadFile({ cloudPath, fileContent: buffer });
	return { success: true, fileID: upload.fileID, fileName };
};
