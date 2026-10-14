const cloud = require('wx-server-sdk');

cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });
const db = cloud.database();
const CATEGORIES = new Set(['formula', 'export', 'layout', 'feature', 'other']);

function clean(value, max) {
	return String(value || '').replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, max);
}

exports.main = async (event) => {
	const category = clean(event.category, 20);
	const content = clean(event.content, 1000);
	const contact = clean(event.contact, 100);
	if (!CATEGORIES.has(category)) throw new Error('无效的反馈分类');
	if (content.length < 5) throw new Error('请至少填写 5 个字的问题描述');

	const context = cloud.getWXContext();
	const date = new Date();
	const ticket = `FB-${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}-${Math.random().toString(36).slice(2, 8).toUpperCase()}`;
	const diagnostics = event.includeDiagnostics ? {
		appVersion: clean(event.appVersion, 40),
		platform: clean(event.platform, 40),
	} : null;

	await db.collection('feedback').add({ data: {
		ticket,
		ownerOpenId: context.OPENID,
		category,
		content,
		contact,
		diagnostics,
		status: 'new',
		createdAt: Date.now(),
	} });
	return { success: true, ticket };
};
