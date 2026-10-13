const storage = require('../../services/storage-service');
const metrics = require('../../services/metrics-service');

function fileName(draft) {
	return `${(draft.title || '文档').replace(/[\\/:*?"<>|]/g, '_')}_${new Date().toISOString().slice(0, 10)}.docx`;
}

Page({
	data: { draft: null, state: 'idle', message: '', fileID: '', name: '' },

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
				},
			})
			.then((response) => {
				const result = response.result || {};
				if (!result.success) {
					throw new Error(result.errMsg || result.message || '生成失败');
				}
				draft.status = 'exported';
				storage.save(draft);
			metrics.track('word_export_succeeded', { draftId: draft.id, attemptId: draft.activeAttemptId });
				this.setData({ state: 'success', fileID: result.fileID, name: result.fileName || this.data.name });
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
				wx.hideLoading();
				wx.openDocument({
					filePath: response.tempFilePath,
					showMenu: true,
					success: () => {
						metrics.track('document_delivery_completed', { draftId: this.data.draft.id, attemptId: this.data.draft.activeAttemptId });
						wx.showToast({ title: '已打开文件', icon: 'success' });
					},
					fail: (error) => wx.showModal({ title: '打开失败', content: error.errMsg || error.message || '请重新下载后再试', showCancel: false }),
				});
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
