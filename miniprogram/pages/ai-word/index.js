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
	const escaped = escapeHtml(text);
	return escaped
		.replace(/`([^`]+)`/g, '<text class="inline-code">$1</text>')
		.replace(/\*\*(.+?)\*\*/g, '<text class="inline-strong">$1</text>')
		.replace(/\*(.+?)\*/g, '<text class="inline-em">$1</text>');
}

function parseBlocks(rawText) {
	const lines = String(rawText || "").replace(/\r\n/g, "\n").split("\n");
	const blocks = [];
	let i = 0;

	while (i < lines.length) {
		const line = lines[i];

		if (!line.trim()) {
			i += 1;
			continue;
		}

		if (line.trim().startsWith("```")) {
			const codeLines = [];
			i += 1;
			while (i < lines.length && !lines[i].trim().startsWith("```")) {
				codeLines.push(lines[i]);
				i += 1;
			}
			if (i < lines.length) {
				i += 1;
			}
			blocks.push({ type: "code", content: codeLines.join("\n") || " " });
			continue;
		}

		if (line.trim().startsWith("$$")) {
			const formulaLines = [];
			i += 1;
			while (i < lines.length && !lines[i].trim().startsWith("$$")) {
				formulaLines.push(lines[i]);
				i += 1;
			}
			if (i < lines.length) {
				i += 1;
			}
			blocks.push({ type: "formula", content: formulaLines.join(" ") || " " });
			continue;
		}

		if (/^#{1,3}\s+/.test(line)) {
			const level = line.match(/^#{1,3}/)[0].length;
			blocks.push({
				type: "heading",
				level,
				content: line.replace(/^#{1,3}\s+/, ""),
			});
			i += 1;
			continue;
		}

		if (/^([-*+]\s+|\d+\.\s+)/.test(line.trim())) {
			const listItems = [];
			while (i < lines.length && /^([-*+]\s+|\d+\.\s+)/.test(lines[i].trim())) {
				listItems.push(lines[i].trim().replace(/^([-*+]\s+|\d+\.\s+)/, ""));
				i += 1;
			}
			blocks.push({ type: "list", content: listItems });
			continue;
		}

		if (line.trim().startsWith(">")) {
			const quoteLines = [];
			while (i < lines.length && lines[i].trim().startsWith(">")) {
				quoteLines.push(lines[i].trim().replace(/^>\s?/, ""));
				i += 1;
			}
			blocks.push({ type: "quote", content: quoteLines.join("\n") });
			continue;
		}

		const paragraphLines = [line.trim()];
		i += 1;
		while (
			i < lines.length &&
			lines[i].trim() &&
			!/^#{1,3}\s+/.test(lines[i]) &&
			!lines[i].trim().startsWith("```") &&
			!lines[i].trim().startsWith("$$") &&
			!/^([-*+]\s+|\d+\.\s+)/.test(lines[i].trim()) &&
			!lines[i].trim().startsWith(">")
		) {
			paragraphLines.push(lines[i].trim());
			i += 1;
		}
		blocks.push({ type: "paragraph", content: paragraphLines.join(" ") });
	}

	return blocks;
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

		const blocks = parseBlocks(text);
		this.setData({
			blocks: blocks.map((item) => ({
				...item,
				rendered: this.getRenderedBlock(item),
			})),
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

	getRenderedBlock(block) {
		if (block.type === "heading") {
			return parseInlineText(block.content);
		}
		if (block.type === "list") {
			return block.content.map((item) => parseInlineText(item));
		}
		return parseInlineText(block.content || "");
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