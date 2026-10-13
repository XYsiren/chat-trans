function issueKey(category, severity, block, message) {
  const range = block && block.sourceRange ? `${block.sourceRange.start}:${block.sourceRange.end}` : 'document';
  return `${category}|${severity}|${range}|${message}`;
}

function make(category, severity, block, message, suggestion, dismissedKeys) {
  const key = issueKey(category, severity, block, message);
  return { id: `issue_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`, key, category, severity, blockId: block && block.id, message, suggestion: suggestion && Object.assign({ confidence: 0.9 }, suggestion), status: (dismissedKeys || []).includes(key) ? 'rejected' : 'pending' };
}

function countUnescaped(source, character) {
  let count = 0;
  for (let i = 0; i < source.length; i += 1) if (source[i] === character && source[i - 1] !== '\\') count += 1;
  return count;
}

function closingSuggestion(source, open, close) {
  const missing = countUnescaped(source, open) - countUnescaped(source, close);
  return missing > 0 ? { before: source, after: `${source}${close.repeat(missing)}`, reason: `补全缺少的 ${close}` } : null;
}

function checkHealth(blocks, template, dismissedKeys) {
  const out = []; let previous = 0; let hasH1 = false;
  const add = (category, severity, block, message, suggestion) => out.push(make(category, severity, block, message, suggestion, dismissedKeys));
  blocks.forEach((block) => {
    if (block.type === 'heading') {
      if (block.level === 1) hasH1 = true;
      if (previous && block.level > previous + 1) add('structure', 'warning', block, '标题层级跳跃。');
      previous = block.level;
    }
    if (block.type === 'formula') {
      const source = String(block.content || '');
      if (!block.closed) add('formula', 'error', block, '公式边界没有闭合。');
      [['(', ')'], ['[', ']'], ['{', '}']].forEach(([open, close]) => {
        if (countUnescaped(source, open) !== countUnescaped(source, close)) add('formula', 'error', block, `${open}${close} 未成对。`, closingSuggestion(source, open, close));
      });
      if (/\\(?:frac|sqrt|text)\s*\{[^}]*$/.test(source)) add('formula', 'warning', block, 'LaTeX 命令参数可能不完整。');
      if (/[，。；]/.test(source)) add('formula', 'warning', block, '公式中包含中文标点，建议确认是否需要规范化。');
    }
    if (typeof block.content === 'string' && /\b(TODO|XXX)\b|待补充/.test(block.content)) add('placeholder', 'warning', block, '发现待补充的占位内容。');
  });
  if (blocks.length && !hasH1) add('structure', 'warning', null, '文档没有一级标题。');
  (template && template.requiredSections || []).forEach((name) => {
    if (!blocks.some((block) => block.type === 'heading' && block.content.indexOf(name) > -1)) add('structure', 'error', null, `“${name}”为该版式的建议必填章节。`);
  });
  return out;
}

function score(items) { return Math.max(0, 100 - items.filter((item) => item.status === 'pending' && item.severity === 'error').length * 15 - items.filter((item) => item.status === 'pending' && item.severity === 'warning').length * 5); }
module.exports = { checkHealth, score };
