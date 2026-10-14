const metrics = require('../../services/metrics-service');

Page({
	data: { category: 'formula', categories: [{ value: 'formula', label: '公式 / 解析' }, { value: 'export', label: 'Word 导出' }, { value: 'layout', label: '排版效果' }, { value: 'feature', label: '功能建议' }, { value: 'other', label: '其他' }], content: '', contact: '', includeDiagnostics: false, submitting: false },
	category(event) { this.setData({ category: event.detail.value }); },
	content(event) { this.setData({ content: event.detail.value }); },
	contact(event) { this.setData({ contact: event.detail.value }); },
	diagnostics(event) { this.setData({ includeDiagnostics: Array.isArray(event.detail.value) && event.detail.value.length > 0 }); },
	submit() {
		const content = this.data.content.trim();
		if (content.length < 5) { wx.showToast({ title: '请至少填写 5 个字', icon: 'none' }); return; }
		if (!wx.cloud) { wx.showToast({ title: '当前无法提交反馈', icon: 'none' }); return; }
		this.setData({ submitting: true });
		const info = wx.getSystemInfoSync();
		wx.cloud.callFunction({ name: 'feedback', data: {
			category: this.data.category, content, contact: this.data.contact,
			includeDiagnostics: this.data.includeDiagnostics,
			appVersion: (wx.getAccountInfoSync().miniProgram || {}).version || 'develop', platform: info.platform || '',
		} }).then((response) => {
			const result = response.result || {};
			if (!result.success) throw new Error(result.errMsg || '提交失败');
			metrics.track('feedback_submitted');
			wx.showModal({ title: '感谢你的反馈', content: `反馈编号：${result.ticket}\n我们会尽快查看。`, showCancel: false, success: () => wx.navigateBack() });
		}).catch((error) => wx.showModal({ title: '提交失败', content: error.errMsg || error.message || '请稍后重试', showCancel: false }))
			.finally(() => this.setData({ submitting: false }));
	},
});
