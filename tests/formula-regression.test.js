const assert = require('assert');
const { parseDocument, normalizeLatex, formulaPreview } = require('../miniprogram/utils/parser');

const cases = [
	{ name: '块级分式', input: '$$\n\\frac{a}{b}\\le t\n$$', types: ['formula'], formula: '\\frac{a}{b}\\le t', closed: true },
	{ name: '单行块公式', input: '$$E=mc^2$$', types: ['formula'], formula: 'E=mc^2', closed: true },
	{ name: '反斜杠块公式', input: '\\[\\sum_{i=1}^{n}i\\]', types: ['formula'], formula: '\\sum_{i=1}^{n}i', closed: true },
	{ name: '独占单美元公式', input: '$x_i^2+y_i^2$', types: ['formula'], formula: 'x_i^2+y_i^2', closed: true },
	{ name: '金额不是公式块', input: '价格为 $10，明天付款。', types: ['paragraph'] },
	{ name: 'math 围栏', input: '```math\n\\int_0^1 x dx\n```', types: ['formula'], formula: '\\int_0^1 x dx', closed: true },
	{ name: '普通代码围栏', input: '```js\nconst x = 1;\n```', types: ['code'] },
	{ name: 'Markdown 表格', input: '| 名称 | 数值 |\n| --- | :---: |\n| alpha | 1 |\n| beta | 2 |', types: ['table'] },
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
assert.strictEqual(formulaPreview('\\boldsymbol{P=\\frac{K-1}{2}}'), 'P=(K-1)/(2)');
assert.strictEqual(formulaPreview('K=3 \\Rightarrow P=(3-1)/2=1'), 'K=3 ⇒ P=(3-1)/2=1');
assert.strictEqual(formulaPreview('5\\times5,\\ K=3,S=1,P=1'), '5×5, K=3,S=1,P=1');
assert.strictEqual(formulaPreview('\\min C=\\sum_{i\\in S_G}(a_iP_{Gi}^2),\\tag{1}'), 'min C=∑₍i∈ S₍G₎₎(aᵢP₍Gi₎²), (1)');
assert.strictEqual(formulaPreview('P_{Gi}-P_{Di}=\\Gamma(V,\\theta,z(G))\\quad(i,j\\in S_B),\\tag{2}'), 'P₍Gi₎-P₍Di₎=Γ(V,θ,z(G)) (i,j∈ S₍B₎), (2)');
assert.strictEqual(formulaPreview('\\underline{P_{Gi}}\\le P_{Gi}\\le\\overline{P_{Gi}}\\quad(i\\in S_G),\\tag{4}'), 'P₍Gi₎≤ P₍Gi₎≤P₍Gi₎ (i∈ S₍G₎), (4)');
assert.strictEqual(formulaPreview('\\mathcal{L}_{\\text{MSE}}=\\frac{1}{N}\\sum_{i=1}^{N}(y_i-\\hat y_i)^2'), 'ℒ₍MSE₎=(1)/(N)∑ᵢ₌₁⁽N⁾(yᵢ-ŷᵢ)²');
const table = parseDocument('| 名称 | 数值 |\n| --- | :---: |\n| alpha | 1 |')[0];
assert.deepStrictEqual(table.content, [['名称', '数值'], ['alpha', '1']]);
assert.strictEqual(table.alignments[1], 'center');
const code = parseDocument('```python\nprint(1)\n```')[0];
assert.strictEqual(code.language, 'python');
console.log(`Formula regression passed: ${cases.length} cases`);
