const { parseDocument, normalizeLatex } = require("../../utils/parser");

const EXAMPLE_TEXT = `# AI 内容格式转换工具

这是一个示例输入，包含标题、列表、代码块和公式。

## 核心能力

- 自动识别标题层级
- 自动拆分段落与列表
- 保留代码块和公式结构

### 代码示例

\`\`\`
function add(a, b) {
  return a + b;
}
\`\`\`

### 公式示例

$$
E = mc^2
$$

> 说明：此页为前端页面骨架，后续可接入云函数生成 Word 文件。
`;

function escapeHtml(source) {
	return String(source)
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/\"/g, "&quot;")
		.replace(/'/g, "&#39;");
}

function parseInlineText(text) {
	const escaped = escapeHtml(normalizeLatex(text));
	return escaped
		.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>')
		.replace(/\*\*(.+?)\*\*/g, '<strong class="inline-strong">$1</strong>')
		.replace(/\*(.+?)\*/g, '<em class="inline-em">$1</em>')
		.replace(/\\\((.+?)\\\)/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`)
		.replace(/\\\[(.+?)\\\]/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`)
		.replace(/\$([^$\n]+)\$/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`);
}

function decorateBlocks(blocks) {
	return blocks.map((block) => {
		if (block.type === "list") {
			return {
				...block,
				renderedItems: (block.content || []).map((item) => parseInlineText(item)),
			};
		}
		if (block.type === "code") {
			return {
				...block,
				renderedText: String(block.content || ""),
			};
		}
		return {
			...block,
			rendered: parseInlineText(block.content || ""),
		};
	});
}

Page({
	data: {
		activeTab: "preview",
		rawText: EXAMPLE_TEXT,
		blocks: [],
		conversionStatus: "idle",
		generatedFileName: "",
		generatedAt: "",
		downloadUrl: "",
		errorMessage: "",
		draftSaved: false,
		dirty: false,
	},

	onLoad() {
		this.setData({ rawText: EXAMPLE_TEXT });
		this.runParse(EXAMPLE_TEXT, true);
	},

	onInputChange(e) {
		const rawText = e.detail.value;
		this.setData({ rawText, dirty: true, draftSaved: false });
	},

	loadExample() {
		this.setData({ rawText: EXAMPLE_TEXT });
		this.runParse(EXAMPLE_TEXT, true);
	},

	clearContent() {
		this.setData({
			rawText: "",
			blocks: [],
			conversionStatus: "idle",
			generatedFileName: "",
			generatedAt: "",
			downloadUrl: "",
			errorMessage: "",
			dirty: false,
		});
	},

	onParseTap() {
		this.runParse(this.data.rawText);
	},

	runParse(rawText, silent) {
		const text = String(rawText || "").trim();
		if (!text) {
			this.setData({
				blocks: [],
				conversionStatus: "error",
				errorMessage: "请输入或粘贴内容后再进行识别。",
				activeTab: "input",
			});
			return;
		}

		const blocks = decorateBlocks(parseDocument(text));
		this.setData({
			blocks,
			conversionStatus: "preview",
			errorMessage: "",
			activeTab: "preview",
			dirty: false,
		});

		if (!silent) {
			wx.showToast({
				title: "结构识别完成",
				icon: "success",
			});
		}
	},

	switchTab(e) {
		const tab = e.currentTarget.dataset.tab;
		this.setData({ activeTab: tab });
	},

	onGenerateWord() {
		if (!this.data.blocks.length) {
			this.setData({
				conversionStatus: "error",
				errorMessage: "请先完成结构识别。",
				activeTab: "input",
			});
			return;
		}

		wx.showLoading({ title: "正在生成 Word..." });
		wx.cloud
			.callFunction({
				name: "quickstartFunctions",
				data: {
					type: "aiWordConversion",
					blocks: this.data.blocks,
				},
			})
			.then((resp) => {
				wx.hideLoading();
				if (resp.result.success) {
					const now = new Date();
					const generatedAt = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
					this.setData({
						conversionStatus: "success",
						generatedFileName: resp.result.fileName,
						generatedAt: generatedAt,
						downloadUrl: resp.result.fileID,
						activeTab: "result",
						errorMessage: "",
					});
					wx.showToast({
						title: "Word 生成成功",
						icon: "success",
					});
				} else {
					this.setData({
						conversionStatus: "error",
						errorMessage: resp.result.errMsg || "Word 生成失败",
						activeTab: "result",
					});
					wx.showToast({
						title: "生成失败",
						icon: "none",
					});
				}
			})
			.catch((err) => {
				wx.hideLoading();
				this.setData({
					conversionStatus: "error",
					errorMessage: err.errMsg || "云函数调用失败",
					activeTab: "result",
				});
				wx.showToast({
					title: "调用失败",
					icon: "none",
				});
			});
	},

	onSaveDraft() {
		this.setData({ draftSaved: true });
		wx.showToast({
			title: "草稿已保存",
			icon: "success",
		});
	},

	onDownload() {
		const fileID = this.data.downloadUrl;
		if (!fileID) {
			wx.showToast({
				title: "文件地址不存在",
				icon: "none",
			});
			return;
		}

		wx.showLoading({ title: "下载中..." });
		wx.cloud
			.downloadFile({
				fileID: fileID,
			})
			.then((res) => {
				wx.hideLoading();
				const tempFilePath = res.tempFilePath;
				wx.saveFile({
					tempFilePath: tempFilePath,
					success(res) {
						wx.showToast({
							title: "已保存到本地",
							icon: "success",
						});
					},
					fail() {
						wx.showModal({
							title: "保存失败",
							content: "请重试",
							showCancel: false,
						});
					},
				});
			})
			.catch(() => {
				wx.hideLoading();
				wx.showToast({
					title: "下载失败",
					icon: "none",
				});
			});
	},
});