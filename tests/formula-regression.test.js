const assert = require('assert');
const { parseDocument, normalizeLatex } = require('../miniprogram/utils/parser');

const cases = [
	{ name: '块级分式', input: '$$\n\\frac{a}{b}\\le t\n$$', types: ['formula'], formula: '\\frac{a}{b}\\le t', closed: true },
	{ name: '单行块公式', input: '$$E=mc^2$$', types: ['formula'], formula: 'E=mc^2', closed: true },
	{ name: '反斜杠块公式', input: '\\[\\sum_{i=1}^{n}i\\]', types: ['formula'], formula: '\\sum_{i=1}^{n}i', closed: true },
	{ name: '独占单美元公式', input: '$x_i^2+y_i^2$', types: ['formula'], formula: 'x_i^2+y_i^2', closed: true },
	{ name: '金额不是公式块', input: '价格为 $10，明天付款。', types: ['paragraph'] },
	{ name: 'math 围栏', input: '```math\n\\int_0^1 x dx\n```', types: ['formula'], formula: '\\int_0^1 x dx', closed: true },
	{ name: '普通代码围栏', input: '```js\nconst x = 1;\n```', types: ['code'] },
	{ name: '行内公式保留在段落', input: '设 $x_i\\in S$，则成立。', types: ['paragraph'] },
	{ name: '未闭合块公式', input: '$$\n\\sqrt{x}', types: ['formula'], formula: '\\sqrt{x}', closed: false },
	{ name: '标题与公式混排', input: '## 定义\n\n\\(L=\\{x\\mid x>0\\}\\)', types: ['heading', 'formula'], formula: 'L=\\{x\\mid x>0\\}', closed: true },
	{ name: '旧式兼容命令', input: '$$\\cfrac{1}{2}+{\\bf x}$$', types: ['formula'], formula: '\\dfrac{1}{2}+{\\mathbf{x}}', closed: true },
	{ name: '集合量词公式', input: '$$\\forall i\\ne j, S_i\\cap S_j=\\emptyset$$', types: ['formula'], formula: '\\forall i\\ne j, S_i\\cap S_j=\\emptyset', closed: true },
];

cases.forEach((test) => {
	const blocks = parseDocument(test.input);
	assert.deepStrictEqual(blocks.map((block) => block.type), test.types, test.name);
	const formula = blocks.find((block) => block.type === 'formula');
	if (test.formula) assert.strictEqual(formula.content, test.formula, test.name);
	if (typeof test.closed === 'boolean') assert.strictEqual(formula.closed, test.closed, test.name);
});

assert.strictEqual(normalizeLatex('\\cfrac{1}{2}+{\\bf x}'), '\\dfrac{1}{2}+{\\mathbf{x}}');
console.log(`Formula regression passed: ${cases.length} cases`);
