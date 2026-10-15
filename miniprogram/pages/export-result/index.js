const storage = require('../../services/storage-service');
const metrics = require('../../services/metrics-service');

function fileName(draft) {
	return `${(draft.title || '文档').replace(/[\\/:*?"<>|]/g, '_')}_${new Date().toISOString().slice(0, 10)}.docx`;
}

Page({
	data: { draft: null, state: 'idle', message: '', fileID: '', name: '', formulaMode: 'compatible' },

	onLoad(query) {
		const draft = storage.get(query.id);
		if (!draft) {
			wx.navigateBack();
			return;
		}
		this.setData({ draft, name: fileName(draft) });
		this.generate();
	},

	generate() {
		const draft = this.data.draft;
		// wx.openDocument 的预览器（包括 Windows/macOS 微信）不会可靠渲染
		// Word 的 OMML 公式对象。统一使用可预览的文本公式，避免桌面端丢公式。
		const formulaMode = 'compatible';
		this.setData({ state: 'generating', message: '正在应用版式…' });

		if (!wx.cloud) {
			this.setData({ state: 'failed', message: '当前环境未启用云开发，无法生成 Word。' });
			return;
		}

		wx.cloud
			.callFunction({
				name: 'generateWord',
				data: {
					draft,
					exportOptions: draft.exportOptions || {},
					formulaMode,
				},
			})
			.then((response) => {
				const result = response.result || {};
				if (!result.success) {
					throw new Error(result.errMsg || result.message || '生成失败');
				}
				if (result.generatorVersion !== 4) {
					throw new Error('云函数尚未更新，请在微信开发者工具中重新部署 generateWord 后再导出。');
				}
				draft.status = 'exported';
				storage.save(draft);
			metrics.track('word_export_succeeded', { draftId: draft.id, attemptId: draft.activeAttemptId });
				this.setData({ state: 'success', fileID: result.fileID, name: result.fileName || this.data.name, formulaMode: result.formulaMode || formulaMode });
			})
			.catch((error) => {
			metrics.track('word_export_failed', { draftId: draft.id, attemptId: draft.activeAttemptId });
				this.setData({ state: 'failed', message: error.errMsg || error.message || '暂时无法生成文档，请检查网络后重试。' });
			});
	},

	download() {
		if (!this.data.fileID) {
			wx.showToast({ title: '文件地址不存在', icon: 'none' });
			return;
		}

		wx.showLoading({ title: '下载中…' });
		wx.cloud
			.getTempFileURL({ fileList: [this.data.fileID] })
			.then((response) => {
				const item = response.fileList && response.fileList[0];
				if (!item || item.status !== 0 || !item.tempFileURL) {
					throw new Error((item && item.errMsg) || '获取下载地址失败');
				}
				return new Promise((resolve, reject) => {
					wx.downloadFile({
						url: item.tempFileURL,
						success: resolve,
						fail: reject,
					});
				});
			})
			.then((response) => {
				if (!response.tempFilePath) {
					throw new Error('文件下载失败');
				}
				return new Promise((resolve, reject) => {
					// 小程序环境不提供系统级“另存为”路径选择器；不再打开微信预览，
					// 直接把文件保存到小程序本地文件区。
					wx.saveFile({
						tempFilePath: response.tempFilePath,
						success: resolve,
						fail: reject,
					});
				});
			})
			.then(() => {
				wx.hideLoading();
				metrics.track('document_delivery_completed', { draftId: this.data.draft.id, attemptId: this.data.draft.activeAttemptId });
				wx.showToast({ title: '文件已保存', icon: 'success' });
			})
			.catch((error) => {
				wx.hideLoading();
				wx.showModal({
					title: '下载失败',
					content: error.errMsg || error.message || '请重试',
					showCancel: false,
				});
			});
	},

	edit() {
		wx.redirectTo({ url: `/pages/editor/index?id=${this.data.draft.id}` });
	},

	retry() {
		this.generate();
	},
});
