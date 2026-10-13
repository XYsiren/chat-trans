const storage=require('./storage-service');
function createDraft(){const now=Date.now();return {id:`draft_${now}_${Math.random().toString(36).slice(2,7)}`,title:'未命名文档',sourceText:'',blocks:[],issues:[],dismissedIssueKeys:[],templateId:'homework',status:'editing',revision:0,createdAt:now,updatedAt:now,changes:[]};}
function saveLocal(draft){draft.updatedAt=Date.now();storage.save(draft);return draft;}
async function saveCloud(draft){saveLocal(draft);if(!wx.cloud)return draft;try{await wx.cloud.callFunction({name:'documentParse',data:{action:'save',draft}});}catch(e){}return draft;}
module.exports={createDraft,saveLocal,saveCloud};
