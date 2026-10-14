const cloud = require('wx-server-sdk');
cloud.init({ env: cloud.DYNAMIC_CURRENT_ENV });
const db = cloud.database();

const ALLOWED = new Set([
	'app_opened',
	'document_processing_started', 'repair_suggestion_shown', 'repair_suggestion_accepted',
	'repair_suggestion_dismissed', 'repair_undone', 'repair_feedback_incorrect',
	'word_export_succeeded', 'word_export_failed', 'document_delivery_completed',
	'feedback_submitted',
]);

exports.main = async (event) => {
	const item = event && event.event;
	if (!item || !ALLOWED.has(item.name) || !item.id) throw new Error('无效的指标事件');
	const context = cloud.getWXContext();
	const data = {
		eventId: String(item.id).slice(0, 80),
		name: item.name,
		at: Number(item.at) || Date.now(),
		attemptId: String(item.attemptId || '').slice(0, 80),
		draftId: String(item.draftId || '').slice(0, 80),
		suggestionId: String(item.suggestionId || '').slice(0, 160),
		repairId: String(item.repairId || '').slice(0, 160),
		formulaCount: Number(item.formulaCount) || 0,
		blockCount: Number(item.blockCount) || 0,
		ownerOpenId: context.OPENID,
	};
	await db.collection('metricsEvents').add({ data });
	return { success: true };
};
