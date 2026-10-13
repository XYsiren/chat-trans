function uniqueBy(events, selector) {
	const map = new Map();
	(events || []).forEach((event) => {
		const key = selector(event);
		if (key) map.set(key, event);
	});
	return map;
}

function percentage(numerator, denominator) {
	return denominator ? Number((numerator / denominator * 100).toFixed(1)) : null;
}

/** 修复建议采纳率：至少被采用一次的建议 / 被展示过的建议。 */
function repairSuggestionAdoption(events) {
	const shown = uniqueBy(events.filter((event) => event.name === 'repair_suggestion_shown'), (event) => event.suggestionId);
	const accepted = uniqueBy(events.filter((event) => event.name === 'repair_suggestion_accepted'), (event) => event.suggestionId);
	const adopted = [...accepted.keys()].filter((id) => shown.has(id)).length;
	return { numerator: adopted, denominator: shown.size, value: percentage(adopted, shown.size) };
}

/** 误修复率：被用户撤销或标记“有问题”的修复 / 已采用修复。 */
function incorrectRepairRate(events) {
	const accepted = uniqueBy(events.filter((event) => event.name === 'repair_suggestion_accepted'), (event) => event.repairId);
	const incorrect = uniqueBy(events.filter((event) => event.name === 'repair_undone' || event.name === 'repair_feedback_incorrect'), (event) => event.repairId);
	const count = [...incorrect.keys()].filter((id) => accepted.has(id)).length;
	return { numerator: count, denominator: accepted.size, value: percentage(count, accepted.size) };
}

/** 端到端交付完成率：开始处理的尝试中，最终成功打开 Word 的比例。 */
function documentDeliveryCompletion(events) {
	const started = uniqueBy(events.filter((event) => event.name === 'document_processing_started'), (event) => event.attemptId);
	const completed = uniqueBy(events.filter((event) => event.name === 'document_delivery_completed'), (event) => event.attemptId);
	const count = [...completed.keys()].filter((id) => started.has(id)).length;
	return { numerator: count, denominator: started.size, value: percentage(count, started.size) };
}

function calculateMetrics(events) {
	return {
		repairSuggestionAdoption: repairSuggestionAdoption(events),
		incorrectRepairRate: incorrectRepairRate(events),
		documentDeliveryCompletion: documentDeliveryCompletion(events),
	};
}

module.exports = { repairSuggestionAdoption, incorrectRepairRate, documentDeliveryCompletion, calculateMetrics };
