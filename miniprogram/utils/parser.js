function makeId() { return `block_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`; }
function make(type, content, start, end, props) { return Object.assign({ id: makeId(), type, content, sourceRange: { start, end } }, props || {}); }
function listLine(line) { return /^\s*(?:[-*+]\s+|\d+[.)]\s+)/.test(line); }
function formulaLine(line) { return /^(?:\$\$[^\n]*\$\$|\\\[[\s\S]*\\\]|\\\([\s\S]*\\\)|\$[^\n$]+\$)$/.test(line.trim()); }
function escapeHtml(source) { return String(source).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
function normalizeLatex(source) {
  return String(source || '')
    .replace(/\\left/g, '')
    .replace(/\\right/g, '')
    .replace(/\\langle/g, '<')
    .replace(/\\rangle/g, '>')
    .replace(/\\lbrace/g, '{')
    .replace(/\\rbrace/g, '}')
    .replace(/\\\{/g, '{')
    .replace(/\\\}/g, '}')
    .replace(/\\dotsc?/g, '…')
    .replace(/\\cdots/g, '…')
    .replace(/\\bigcup/g, '∪')
    .replace(/\\cup/g, '∪')
    .replace(/\\cap/g, '∩')
    .replace(/\\subseteq/g, '⊆')
    .replace(/\\subset/g, '⊂')
    .replace(/\\supseteq/g, '⊇')
    .replace(/\\supset/g, '⊃')
    .replace(/\\emptyset/g, '∅')
    .replace(/\\exists/g, '∃')
    .replace(/\\forall/g, '∀')
    .replace(/\\in/g, '∈')
    .replace(/\\notin/g, '∉')
    .replace(/\\leq/g, '≤')
    .replace(/\\le/g, '≤')
    .replace(/\\geq/g, '≥')
    .replace(/\\ge/g, '≥')
    .replace(/\\ne/g, '≠')
    .replace(/\\approx/g, '≈')
    .replace(/\\times/g, '×')
    .replace(/\\bigg\|/g, '|')
    .replace(/\\big\|/g, '|')
    .replace(/\\vert/g, '|')
    .replace(/\\sum/g, 'Σ')
    .replace(/\\prod/g, 'Π')
    .replace(/\\frac/g, 'frac')
    .replace(/\\sqrt/g, 'sqrt')
    .replace(/\\text\{([^{}]*)\}/g, '$1')
    .replace(/\\mathrm\{([^{}]*)\}/g, '$1')
    .replace(/\\mathbf\{([^{}]*)\}/g, '$1')
    .replace(/\\mathit\{([^{}]*)\}/g, '$1')
    .replace(/\\operatorname\{([^{}]*)\}/g, '$1')
    .replace(/\\,/g, ' ')
    .replace(/\\!/g, '')
    .replace(/\\_/g, '_')
    .replace(/\\%/g, '%')
    .replace(/\\&/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}
function renderInline(source) {
  const escaped = escapeHtml(normalizeLatex(source));
  return escaped
    .replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong class="inline-strong">$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em class="inline-em">$1</em>')
    .replace(/\\\((.+?)\\\)/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`)
    .replace(/\\\[(.+?)\\\]/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`)
    .replace(/\$([^$\n]+)\$/g, (_, expr) => `<span class="inline-formula">${escapeHtml(normalizeLatex(expr))}</span>`);
}
function parseDocument(source) {
  const lines = String(source || '').replace(/\r\n?/g, '\n').split('\n'); const blocks = []; let i = 0; let pos = 0;
  while (i < lines.length) {
    const line = lines[i], start = pos; pos += line.length + 1;
    if (!line.trim()) { i++; continue; }
    if (/^\s*```/.test(line)) { const rows=[]; i++; while(i<lines.length&&!/^\s*```/.test(lines[i])) { rows.push(lines[i]); pos+=lines[i].length+1; i++; } if(i<lines.length) {pos+=lines[i].length+1;i++;} blocks.push(make('code',rows.join('\n'),start,pos)); continue; }
    const trimmed = line.trim();
    if (/^(\$\$|\\\[|\\\()/.test(trimmed) || formulaLine(trimmed)) {
      const open = trimmed.startsWith('$$') ? '$$' : trimmed.startsWith('\\[') ? '\\[' : trimmed.startsWith('\\(') ? '\\(' : null;
      if (open) {
        const close = open === '$$' ? '$$' : open === '\\[' ? '\\]' : '\\)';
        const firstRow = trimmed.replace(new RegExp(`^${open === '$$' ? '\\\$\\\$' : open === '\\\\[' ? '\\\\\\[' : '\\\\\\('}`), '').replace(new RegExp(`(${close === '$$' ? '\\\$\\\$' : close === '\\\\]' ? '\\\\\\]' : '\\\\\\)'})$`), '');
        const rows = firstRow ? [firstRow] : [];
        let closed = trimmed.endsWith(close) || trimmed.endsWith(close.replace('\\', ''));
        if (!closed) {
          i++;
          while (i < lines.length) {
            const row = lines[i];
            pos += row.length + 1;
            if (row.trim() === close) { closed = true; i++; break; }
            rows.push(row);
            i++;
          }
        } else {
          i++;
        }
        const formulaContent = rows.join('\n').trim();
        const displayText = normalizeLatex(formulaContent);
        blocks.push(make('formula', formulaContent, start, pos, { formulaFormat: 'latex', closed, rendered: `<span class="inline-formula">${escapeHtml(displayText)}</span>`, displayText }));
        continue;
      }
    }
    const heading = line.match(/^(#{1,6})\s+(.+)$/); if(heading) {blocks.push(make('heading',heading[2],start,pos,{level:heading[1].length,rendered:renderInline(heading[2])}));i++;continue;}
    if(/^\s*>/.test(line)) {const rows=[line.replace(/^\s*>\s?/, '')];i++;while(i<lines.length&&/^\s*>/.test(lines[i])){rows.push(lines[i].replace(/^\s*>\s?/,''));pos+=lines[i].length+1;i++;}const content=rows.join('\n');blocks.push(make('quote',content,start,pos,{rendered:renderInline(content)}));continue;}
    if(listLine(line)) {const rows=[];const renderedItems=[];const ordered=/^\s*\d+[.)]\s+/.test(line);let row=line;while(row&&listLine(row)){const item=row.replace(/^\s*(?:[-*+]\s+|\d+[.)]\s+)/,'');rows.push(item);renderedItems.push(renderInline(item));i++;if(i>=lines.length)break;row=lines[i];if(listLine(row))pos+=row.length+1;}blocks.push(make('list',rows,start,pos,{ordered,renderedItems}));continue;}
    const rows=[line.trim()];i++;while(i<lines.length&&lines[i].trim()&&!/^\s*```/.test(lines[i])&&!/^(\$\$|\\\[|\\\()/.test(lines[i].trim())&&!formulaLine(lines[i].trim())&&!/^#{1,6}\s+/.test(lines[i])&&!/^\s*>/.test(lines[i])&&!listLine(lines[i])){rows.push(lines[i].trim());pos+=lines[i].length+1;i++;}const content=rows.join('\n');blocks.push(make('paragraph',content,start,pos,{rendered:renderInline(content)}));
  } return blocks;
}
module.exports = { parseDocument, renderInline, normalizeLatex };
