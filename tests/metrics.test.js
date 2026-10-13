const assert = require('assert');
const { calculateMetrics } = require('../miniprogram/utils/metrics');

const events = [
	{ name: 'repair_suggestion_shown', suggestionId: 's1' }, { name: 'repair_suggestion_shown', suggestionId: 's2' },
	{ name: 'repair_suggestion_accepted', suggestionId: 's1', repairId: 'r1' },
	{ name: 'repair_undone', repairId: 'r1' },
	{ name: 'document_processing_started', attemptId: 'a1' }, { name: 'document_processing_started', attemptId: 'a2' },
	{ name: 'document_delivery_completed', attemptId: 'a1' },
];
const result = calculateMetrics(events);
assert.deepStrictEqual(result.repairSuggestionAdoption, { numerator: 1, denominator: 2, value: 50 });
assert.deepStrictEqual(result.incorrectRepairRate, { numerator: 1, denominator: 1, value: 100 });
assert.deepStrictEqual(result.documentDeliveryCompletion, { numerator: 1, denominator: 2, value: 50 });
console.log('Metric calculations passed');
