function makeId() {
  return `block_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function make(type, content, start, end, props) {
  return Object.assign({ id: makeId(), type, content, sourceRange: { start, end } }, props || {});
}

function escapeHtml(source) {
  return String(source == null ? '' : source)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function listLine(line) { return /^\s*(?:[-*+]\s+|\d+[.)]\s+)/.test(line); }
function isMathFence(line) { return /^\s*```(?:math|latex|tex)\s*$/i.test(line); }
function isCodeFence(line) { return /^\s*```/.test(line); }

function isStandaloneSingleDollar(line) {
  const source = line.trim();
  if (!/^\$(?!\$).+\$$/.test(source) || /\$\$$/.test(source)) return false;
  // PasteMD 的 fix_single_dollar_block 思路：独占一行的 $...$ 当作块公式；排除金额。
  const expr = source.slice(1, -1).trim();
  return !!expr && !/^\d+(?:[,.]\d{1,2})?$/.test(expr);
}

/**
 * 只做确定、等价的兼容性替换。不得将 LaTeX 命令替换成普通字符，
 * 否则会破坏后续 MathML/OMML 转换所需的语义。
 */
function normalizeLatex(source) {
  return String(source == null ? '' : source)
    .replace(/\r\n?/g, '\n')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\\cfrac\b/g, '\\dfrac')
    .replace(/\\mathnormal\b/g, '\\mathit')
    .replace(/\{\\kern\s*[^}]+\}/g, '\\qquad')
    .replace(/\\kern\s*-?\d*\.?\d+(?:pt|em|cm|mm|ex|bp)\b/g, '\\qquad')
    .replace(/\{\s*\\cal\s+([^{}]+?)\s*\}/g, '{\\mathcal{$1}}')
    .replace(/\{\s*\\bf\s+([^{}]+?)\s*\}/g, '{\\mathbf{$1}}')
    .replace(/\{\s*\\it\s+([^{}]+?)\s*\}/g, '{\\mathit{$1}}')
    .replace(/\\cal\s+([A-Za-z])/g, '\\mathcal{$1}')
    .replace(/\\bf\s+([A-Za-z])/g, '\\mathbf{$1}')
    .replace(/\\it\s+([A-Za-z])/g, '\\mathit{$1}')
    .trim();
}

function formulaPreview(latex) {
  // 当前未接入 TeX 渲染器，预览保留原始 LaTeX，绝不以“好看”为由丢失语义。
  return normalizeLatex(latex).replace(/\s*\n\s*/g, ' ');
}

function renderInline(source) {
  const text = String(source == null ? '' : source);
  const token = /(`[^`\n]+`|\*\*[^*\n]+\*\*|\*[^*\n]+\*|\\\([\s\S]+?\\\)|\\\[[\s\S]+?\\\]|\$\$[\s\S]+?\$\$|\$(?!\s)(?:\\.|[^$\n])+?\$)/g;
  let result = ''; let cursor = 0; let match;
  while ((match = token.exec(text))) {
    result += escapeHtml(text.slice(cursor, match.index));
    const value = match[0];
    if (value.startsWith('`')) result += `<code class="inline-code">${escapeHtml(value.slice(1, -1))}</code>`;
    else if (value.startsWith('**')) result += `<strong class="inline-strong">${escapeHtml(value.slice(2, -2))}</strong>`;
    else if (value.startsWith('*')) result += `<em class="inline-em">${escapeHtml(value.slice(1, -1))}</em>`;
    else {
      const latex = value.startsWith('\\(') || value.startsWith('\\[') ? value.slice(2, -2)
        : value.startsWith('$$') ? value.slice(2, -2) : value.slice(1, -1);
      result += `<span class="inline-formula">${escapeHtml(formulaPreview(latex))}</span>`;
    }
    cursor = token.lastIndex;
  }
  return result + escapeHtml(text.slice(cursor));
}

function formulaBlock(content, start, end, closed, delimiter) {
  const latex = normalizeLatex(content);
  return make('formula', latex, start, end, { formulaFormat: 'latex', closed, delimiter, displayText: formulaPreview(latex) });
}

function parseDocument(source) {
  const input = String(source == null ? '' : source).replace(/\r\n?/g, '\n');
  const lines = input.split('\n');
  const offsets = []; let offset = 0;
  lines.forEach((line) => { offsets.push(offset); offset += line.length + 1; });
  const blocks = []; let i = 0;
  const blockEnd = (lineIndex) => offsets[lineIndex] + lines[lineIndex].length + (lineIndex < lines.length - 1 ? 1 : 0);

  while (i < lines.length) {
    const line = lines[i]; const trimmed = line.trim(); const start = offsets[i];
    if (!trimmed) { i += 1; continue; }

    // AI 平台经常以 ```math 包装公式；只转换显式 math fence，普通代码不受影响。
    if (isMathFence(line)) {
      const rows = []; const first = i; let closed = false; i += 1;
      while (i < lines.length) {
        if (isCodeFence(lines[i])) { closed = true; i += 1; break; }
        rows.push(lines[i]); i += 1;
      }
      blocks.push(formulaBlock(rows.join('\n'), start, blockEnd(Math.max(first, i - 1)), closed, '```math'));
      continue;
    }

    if (isCodeFence(line)) {
      const rows = []; const first = i; i += 1;
      while (i < lines.length && !isCodeFence(lines[i])) { rows.push(lines[i]); i += 1; }
      const closed = i < lines.length; if (closed) i += 1;
      blocks.push(make('code', rows.join('\n'), start, blockEnd(Math.max(first, i - 1)), { closed }));
      continue;
    }

    const delimiters = trimmed.startsWith('$$') ? ['$$', '$$'] : trimmed.startsWith('\\[') ? ['\\[', '\\]']
      : trimmed.startsWith('\\(') ? ['\\(', '\\)'] : null;
    if (delimiters) {
      const open = delimiters[0]; const close = delimiters[1]; const afterOpen = trimmed.slice(open.length); const closeAt = afterOpen.indexOf(close);
      // 同一行的分隔符只有在末尾没有正文时，才归为独立公式块，避免吞掉文本。
      if (closeAt >= 0 && !afterOpen.slice(closeAt + close.length).trim()) {
        blocks.push(formulaBlock(afterOpen.slice(0, closeAt), start, blockEnd(i), true, open)); i += 1; continue;
      }
      const rows = afterOpen ? [afterOpen] : []; const first = i; let closed = false; i += 1;
      while (i < lines.length) {
        const closeIndex = lines[i].indexOf(close);
        if (closeIndex >= 0) {
          rows.push(lines[i].slice(0, closeIndex));
          closed = !lines[i].slice(closeIndex + close.length).trim(); i += 1; break;
        }
        rows.push(lines[i]); i += 1;
      }
      blocks.push(formulaBlock(rows.join('\n'), start, blockEnd(Math.max(first, i - 1)), closed, open));
      continue;
    }

    if (isStandaloneSingleDollar(line)) {
      blocks.push(formulaBlock(trimmed.slice(1, -1), start, blockEnd(i), true, '$')); i += 1; continue;
    }

    const heading = line.match(/^(#{1,6})\s+(.+)$/);
    if (heading) {
      blocks.push(make('heading', heading[2], start, blockEnd(i), { level: heading[1].length, rendered: renderInline(heading[2]) })); i += 1; continue;
    }

    if (/^\s*>/.test(line)) {
      const rows = []; const first = i;
      while (i < lines.length && /^\s*>/.test(lines[i])) { rows.push(lines[i].replace(/^\s*>\s?/, '')); i += 1; }
      const content = rows.join('\n'); blocks.push(make('quote', content, start, blockEnd(i - 1), { rendered: renderInline(content) })); continue;
    }

    if (listLine(line)) {
      const rows = []; const renderedItems = []; const ordered = /^\s*\d+[.)]\s+/.test(line); const first = i;
      while (i < lines.length && listLine(lines[i])) {
        const item = lines[i].replace(/^\s*(?:[-*+]\s+|\d+[.)]\s+)/, ''); rows.push(item); renderedItems.push(renderInline(item)); i += 1;
      }
      blocks.push(make('list', rows, start, blockEnd(i - 1), { ordered, renderedItems })); continue;
    }

    const rows = [line.trim()]; const first = i; i += 1;
    while (i < lines.length) {
      const next = lines[i];
      if (!next.trim() || isCodeFence(next) || isMathFence(next) || /^(\$\$|\\\[|\\\()/.test(next.trim()) || isStandaloneSingleDollar(next) || /^#{1,6}\s+/.test(next) || /^\s*>/.test(next) || listLine(next)) break;
      rows.push(next.trim()); i += 1;
    }
    const content = rows.join('\n'); blocks.push(make('paragraph', content, start, blockEnd(i - 1), { rendered: renderInline(content) }));
  }
  return blocks;
}

module.exports = { parseDocument, renderInline, normalizeLatex, formulaPreview };
