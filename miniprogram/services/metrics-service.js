const LOCAL_KEY = 'document_metric_events_v1';

function eventId() { return `metric_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`; }
function localEvents() { return wx.getStorageSync(LOCAL_KEY) || []; }

function track(name, payload) {
	const event = Object.assign({ id: eventId(), name, at: Date.now() }, payload || {});
	// 只记录匿名业务元数据，严禁写入 sourceText、公式内容和用户个人信息。
	const events = localEvents();
	events.push(event);
	wx.setStorageSync(LOCAL_KEY, events.slice(-1000));
	if (wx.cloud) wx.cloud.callFunction({ name: 'metrics', data: { event } }).catch(() => {});
	return event;
}

function listLocal() { return localEvents(); }
module.exports = { track, listLocal };
