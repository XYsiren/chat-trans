const katex = require('@rojer/katex-mini');
const parseLatex = katex.default || katex.parseLatex || katex;

Component({
	properties: {
		block: {
			type: Object,
			observer(block) {
				this.renderFormula(block);
			},
		},
		issueCount: Number,
		active: Boolean,
	},

	data: { formulaNodes: [], formulaError: false },

	methods: {
		renderFormula(block) {
			if (!block || block.type !== 'formula') {
				this.setData({ formulaNodes: [], formulaError: false });
				return;
			}

			try {
				const formulaNodes = parseLatex(block.content || '', {
					displayMode: true,
					throwError: true,
					strict: 'ignore',
				});
				this.setData({ formulaNodes, formulaError: false });
			} catch (error) {
				// 个别不完整或非标准 LaTeX 仍保留可读文本，不能让整页预览中断。
				this.setData({ formulaNodes: [], formulaError: true });
			}
		},

		tap() {
			this.triggerEvent('tap', { blockId: this.data.block.id });
		},
	},
});
