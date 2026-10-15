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
		// 微信手机文档预览器不渲染 Word 原生 OMML；手机端改用兼容公式文本，
		// Windows/macOS 端仍导出可编辑 OMML 公式。
		const system = wx.getSystemInfoSync();
		const isDesktop = system.platform === 'windows' || system.platform === 'mac';
		const formulaMode = isDesktop ? 'editable' : 'compatible';
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
				if (!response.tempFilePath) throw new Error('文件下载失败');
				const savedPath = `${wx.env.USER_DATA_PATH}/${this.data.name}`;
				return new Promise((resolve) => {
					wx.getFileSystemManager().saveFile({
						tempFilePath: response.tempFilePath,
						filePath: savedPath,
						success: (result) => resolve(result.savedFilePath || savedPath),
						fail: () => resolve(response.tempFilePath),
					});
				});
			})
			.then((filePath) => {
				wx.hideLoading();
				const system = wx.getSystemInfoSync();
				const isDesktop = system.platform === 'windows' || system.platform === 'mac';
				if (isDesktop && typeof wx.saveFileToDisk === 'function') {
					// PC 微信会弹出系统“另存为”窗口，由用户选择磁盘位置和文件名。
					wx.saveFileToDisk({
						filePath,
						success: () => {
							metrics.track('document_delivery_completed', { draftId: this.data.draft.id, attemptId: this.data.draft.activeAttemptId });
							wx.showToast({ title: '已导出到所选位置', icon: 'success' });
						},
						fail: (error) => {
							if (/cancel/i.test(error.errMsg || '')) return;
							wx.showModal({ title: '导出失败', content: error.errMsg || error.message || '请重新导出', showCancel: false });
						},
					});
					return;
				}

				// 手机端不支持系统路径选择，打开微信文件页后可通过右上角菜单转发或保存。
				wx.openDocument({
					filePath,
					showMenu: true,
					success: () => metrics.track('document_delivery_completed', { draftId: this.data.draft.id, attemptId: this.data.draft.activeAttemptId }),
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
