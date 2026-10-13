const storage = require('../../services/storage-service');
const { createDraft, saveLocal, saveCloud } = require('../../services/document-service');
const { parseDocument } = require('../../utils/parser');
const { checkHealth, score } = require('../../utils/health');
const { TEMPLATES, MAX_SOURCE_LENGTH } = require('../../utils/constants');

const example = '# 文档标题\n\n这里输入正文内容。支持 Markdown 标题、列表、代码块与 LaTeX 公式。\n\n## 公式示例\n\n$$\nE = mc^2\n$$\n\n- 第一项\n- 第二项';

Page({
  data: {
    draft: null,
    tab: 'edit',
    sourceDirty: false,
    panel: false,
    score: 100,
    errors: 0,
    warnings: 0,
    max: MAX_SOURCE_LENGTH,
    visibleIssues: [],
    activeBlockId: '',
  },

  onLoad(query) {
    let draft = query.id && storage.get(query.id);
    if (!draft) {
      draft = saveLocal(createDraft());
    }
    this.load(draft);
  },

  load(draft) {
    this.setData({ draft, sourceDirty: false });
    this.refresh(draft);
  },

  refresh(draft) {
    const issues = checkHealth(draft.blocks || [], TEMPLATES.find((template) => template.id === draft.templateId), draft.dismissedIssueKeys || []);
    draft.issues = issues;
    const pendingIssues = issues.filter((item) => item.status === 'pending');
    this.setData({
      'draft.issues': issues,
      visibleIssues: pendingIssues,
      score: score(issues),
      errors: pendingIssues.filter((item) => item.severity === 'error').length,
      warnings: pendingIssues.filter((item) => item.severity === 'warning').length,
    });
    saveLocal(draft);
  },

  input(event) {
    const value = event.detail.value;
    this.setData({ 'draft.sourceText': value, sourceDirty: true });
    clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => saveLocal(this.data.draft), 800);
  },

  title(event) {
    this.setData({ 'draft.title': event.detail.value || '未命名文档' });
  },

  save() {
    saveCloud(this.data.draft).then(() => wx.showToast({ title: '草稿已保存', icon: 'success' }));
  },

  example() {
    this.setData({ 'draft.sourceText': example, sourceDirty: true });
  },

  clear() {
    wx.showModal({
      title: '清空内容',
      content: '确定清空当前输入内容吗？',
      success: (result) => {
        if (result.confirm) {
          this.setData({ 'draft.sourceText': '', sourceDirty: true });
        }
      },
    });
  },

  parse() {
    const text = (this.data.draft.sourceText || '').trim();
    if (!text) {
      wx.showToast({ title: '请先填写内容', icon: 'none' });
      return;
    }

    const draft = this.data.draft;
    draft.blocks = parseDocument(text);
    draft.revision++;
    draft.status = 'editing';
    this.setData({ draft, sourceDirty: false, tab: 'preview' });
    this.refresh(draft);
    wx.showToast({ title: '结构识别完成', icon: 'success' });
  },

  switch(event) {
    this.setData({ tab: event.currentTarget.dataset.tab });
  },

  issues() {
    if (!this.data.draft.blocks.length) {
      wx.showToast({ title: '请先识别内容', icon: 'none' });
      return;
    }
    this.setData({ panel: true });
  },

  close() {
    this.setData({ panel: false });
  },

  noop() {},

  next() {
    if (this.data.sourceDirty || !this.data.draft.blocks.length) return;
    saveLocal(this.data.draft);
    wx.navigateTo({ url: `/pages/template-select/index?id=${this.data.draft.id}` });
  },

  scrollToBlock(blockId) {
    setTimeout(() => {
      wx.pageScrollTo({ selector: `#${blockId}`, duration: 250, offsetTop: 80 });
    }, 80);
  },

  focus(event) {
    const block = this.data.draft.blocks.find((item) => item.id === event.detail.blockId);
    if (!block) return;
    this.setData({ tab: 'edit', activeBlockId: block.id });
    this.scrollToBlock(block.id);
    wx.showToast({ title: '已定位到对应内容', icon: 'none' });
  },

  focusIssue(event) {
    const issue = this.data.draft.issues.find((item) => item.id === event.detail.issueId);
    if (!issue) return;

    this.setData({ panel: false });
    if (issue.blockId) {
      this.setData({ tab: 'edit', activeBlockId: issue.blockId });
      this.scrollToBlock(issue.blockId);
      wx.showToast({ title: '已定位到对应内容', icon: 'none' });
    } else {
      wx.showToast({ title: '该提醒没有对应块', icon: 'none' });
    }
  },

  reject(event) {
    const issue = this.data.draft.issues.find((item) => item.id === event.detail.issueId);
    if (!issue) return;
    const keys = this.data.draft.dismissedIssueKeys || [];
    if (!keys.includes(issue.key)) keys.push(issue.key);
    this.data.draft.dismissedIssueKeys = keys;
    issue.status = 'rejected';
    this.refresh(this.data.draft);
    this.setData({ panel: true });
    wx.showToast({ title: '本次内容不再提示', icon: 'none' });
  },

  accept(event) {
    const issue = this.data.draft.issues.find((item) => item.id === event.detail.issueId);
    if (!issue || !issue.suggestion) return;

    const block = this.data.draft.blocks.find((item) => item.id === issue.blockId);
    if (!block || block.content !== issue.suggestion.before) return;

    const range = block.sourceRange || {};
    const source = this.data.draft.sourceText || '';
    const part = source.slice(range.start, range.end);
    const index = part.indexOf(issue.suggestion.before);
    if (index < 0) {
      wx.showToast({ title: '内容已变化，请重新识别', icon: 'none' });
      return;
    }

    this.data.draft.sourceText = `${source.slice(0, range.start)}${part.slice(0, index)}${issue.suggestion.after}${part.slice(index + issue.suggestion.before.length)}${source.slice(range.end)}`;
    this.data.draft.blocks = parseDocument(this.data.draft.sourceText);
    this.data.draft.revision++;
    this.refresh(this.data.draft);
    this.setData({ draft: this.data.draft, panel: true, sourceDirty: false });
    wx.showToast({ title: '已完成修复', icon: 'success' });
  },
});
