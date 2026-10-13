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
const { Math: WordMath, latexToMath } = require('docx/math');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });

const headingLevels = [
	HeadingLevel.HEADING_1,
	HeadingLevel.HEADING_2,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
	HeadingLevel.HEADING_3,
];

function mathRun(latex) {
	try {
		return new WordMath({ children: latexToMath(String(latex || '').trim()) });
	} catch (error) {
		// 不把失败公式伪装成正常排版；保留原始源码供用户在 Word 中核对。
		return new TextRun({ text: `[公式未转换: ${String(latex || '').trim()}]`, italics: true, color: 'B64035' });
	}
}

function inlineRuns(text) {
	const source = String(text || '').replace(/\r\n?/g, ' ');
	const runs = [];
	const pattern = /(`[^`]+`|\*\*[^*]+\*\*|\*[^*]+\*|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$\$[\s\S]+?\$\$|\$[^$\n]+?\$)/g;
	let lastIndex = 0;
	let match = pattern.exec(source);
	while (match) {
		if (match.index > lastIndex) {
			runs.push(new TextRun({ text: source.slice(lastIndex, match.index) }));
		}

		const token = match[0];
		if (token.startsWith('**')) {
			runs.push(new TextRun({ text: token.slice(2, -2), bold: true }));
		} else if (token.startsWith('*')) {
			runs.push(new TextRun({ text: token.slice(1, -1), italics: true }));
		} else if (token.startsWith('`')) {
			runs.push(new TextRun({ text: token.slice(1, -1), font: 'Courier New' }));
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
		runs.push(new TextRun({ text: source.slice(lastIndex) }));
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
