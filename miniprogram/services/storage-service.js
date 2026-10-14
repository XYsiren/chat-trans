const KEY='document_drafts_v1';
function all(){return wx.getStorageSync(KEY)||[];} function list(){return all().sort((a,b)=>b.updatedAt-a.updatedAt);} function get(id){return all().find(x=>x.id===id)||null;} function save(d){const a=all(),i=a.findIndex(x=>x.id===d.id);if(i<0)a.unshift(d);else a.splice(i,1,d);wx.setStorageSync(KEY,a);return d;} function remove(id){wx.setStorageSync(KEY,all().filter(x=>x.id!==id));} function clear(){wx.removeStorageSync(KEY);}
module.exports={list,get,save,remove,clear};
