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
function isTableRow(line) { return /^\s*\|?.+\|.+\|?\s*$/.test(line); }
function isTableSeparator(line) {
  if (!isTableRow(line)) return false;
  const cells = tableCells(line);
  return cells.length > 1 && cells.every((cell) => /^:?-{3,}:?$/.test(cell.trim()));
}
function tableCells(line) {
  const source = String(line || '').trim().replace(/^\||\|$/g, '');
  const cells = []; let cell = ''; let escaped = false;
  for (const char of source) {
    if (escaped) { cell += char; escaped = false; }
    else if (char === '\\') escaped = true;
    else if (char === '|') { cells.push(cell.trim()); cell = ''; }
    else cell += char;
  }
  cells.push(cell.trim());
  return cells;
}

function isStandaloneSingleDollar(line) {
  const source = line.trim();
  if (!/^\$(?!\$).+\$$/.test(source) || /\$\$$/.test(source)) return false;
  // PasteMD 的 fix_single_dollar_block 思路：独占一行的 $...$ 当作块公式；排除金额。
  const expr = source.slice(1, -1).trim();
  return !!expr && !/^\d+(?:[,.]\d{1,2})?$/.test(expr);
}

function bareMathContent(line) {
  const source = line.trim().replace(/[；。]$/, '').trim();
  if (!source || /[\u4e00-\u9fff]/.test(source)) return null;
  const hasLatex = /\\(?:frac|dfrac|tfrac|sqrt|sum|prod|int|times|cdot|leq?|geq?|Rightarrow|rightarrow|leftarrow|boldsymbol|mathbf|mathit|mathrm|in|cap|cup|;)/.test(source);
  const looksLikeEquation = /[=_^]/.test(source) && /^[A-Za-z0-9_{}^\\.,;:+\-*/=<>()[\]\s]+$/.test(source);
  return hasLatex || looksLikeEquation ? source : null;
}

function wrapBareInlineVariables(source) {
  const text = String(source || '');
  // 含已有数学分隔符时不再二次包裹，避免破坏用户写好的公式。
  if (/\\\(|\\\[|\$/.test(text)) return text;
  return text.replace(/\b(?:[A-Za-z]_\{[A-Za-z][A-Za-z0-9]*\}\s*=\s*[A-Za-z]_\{[A-Za-z][A-Za-z0-9]*\}|[A-Za-z]_\{[A-Za-z][A-Za-z0-9]*\})/g, (match) => `\\(${match}\\)`);
}

/**
 * 只做确定、等价的兼容性替换。不得将 LaTeX 命令替换成普通字符，
 * 否则会破坏后续 MathML/OMML 转换所需的语义。
 */
function normalizeLatex(source) {
  return String(source == null ? '' : source)
    .replace(/\r\n?/g, '\n')
		// 聊天窗口常把 LaTeX 命令转义为 `\\frac`、`\\lambda`；数组的 `\\` 换行
		// 后面通常不是字母，因此不会被这一规则误改。
		.replace(/\\{2,}(?=[A-Za-z])/g, '\\')
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
		// 有些来源会写成 `$\\(x\\)$` 或 `\\boldsymbol{\\(x\\)}`：外层已经
		// 是数学环境，内层 `\\(` / `\\)` 不是 LaTeX 内容，传给 KaTeX/Word 会报错。
		// 成对移除即可同时保留逗号、空格和其他真正的公式语义。
		.replace(/\\\(([\s\S]*?)\\\)/g, '$1')
    .trim();
}

function normalizeMathDelimiters(source) {
  // 从聊天窗口复制的 `\\(`、`\\[` 不是公式内容的一部分，而是被多转义了一次的分隔符。
	return String(source == null ? '' : source).replace(/\\{2,}(?=[\[\]()])/g, '\\');
}

function formulaPreview(latex) {
  // `content` 始终保留 LaTeX；这里只生成小程序内的可读预览文本。
  // 小程序原生 rich-text 不具备 TeX 排版能力，不能把 LaTeX 原样直接展示给用户。
  let value = normalizeLatex(latex).replace(/\s*\n\s*/g, ' ');

  // 先展开最内层的分式和文字命令，循环可处理有限层嵌套。
  for (let i = 0; i < 6; i += 1) {
    const next = value
      // 分式内经常出现 H_{in} 这类一层嵌套的下标，不能只匹配纯文本参数。
      .replace(/\\(?:dfrac|tfrac|cfrac|frac)\{((?:[^{}]|\{[^{}]*\})*)\}\{((?:[^{}]|\{[^{}]*\})*)\}/g, '($1)/($2)')
      .replace(/\\sqrt\{([^{}]*)\}/g, '√($1)')
      .replace(/\\mathcal\{L\}/g, 'ℒ')
      .replace(/\\(?:text|mathrm|mathbf|mathit|mathcal|boldsymbol|operatorname)\{([^{}]*)\}/g, '$1');
    if (next === value) break;
    value = next;
  }

  const symbols = [
    ['\\\\left', ''], ['\\\\right', ''], ['\\\\bigg', ''], ['\\\\Bigg', ''], ['\\\\big', ''], ['\\\\Big', ''],
    ['\\\\langle', '⟨'], ['\\\\rangle', '⟩'], ['\\\\lceil', '⌈'], ['\\\\rceil', '⌉'], ['\\\\lfloor', '⌊'], ['\\\\rfloor', '⌋'],
    ['\\\\bigcup', '⋃'], ['\\\\bigcap', '⋂'], ['\\\\sum', '∑'], ['\\\\prod', '∏'], ['\\\\int', '∫'],
    ['\\\\exists', '∃'], ['\\\\forall', '∀'], ['\\\\emptyset', '∅'], ['\\\\infty', '∞'], ['\\\\partial', '∂'],
    ['\\\\subseteq', '⊆'], ['\\\\supseteq', '⊇'], ['\\\\subset', '⊂'], ['\\\\supset', '⊃'], ['\\\\notin', '∉'], ['\\\\in', '∈'],
    ['\\\\leq', '≤'], ['\\\\geq', '≥'], ['\\\\le', '≤'], ['\\\\ge', '≥'], ['\\\\neq', '≠'], ['\\\\ne', '≠'], ['\\\\approx', '≈'],
    ['\\\\cdots', '⋯'], ['\\\\ldots', '…'], ['\\\\dots', '…'], ['\\\\times', '×'], ['\\\\cdot', '·'], ['\\\\cap', '∩'], ['\\\\cup', '∪'],
    ['\\\\to', '→'], ['\\\\rightarrow', '→'], ['\\\\leftarrow', '←'], ['\\\\mid', '|'], ['\\\\vert', '|'], ['\\\\|', '|'],
    ['\\\\alpha', 'α'], ['\\\\beta', 'β'], ['\\\\gamma', 'γ'], ['\\\\delta', 'δ'], ['\\\\lambda', 'λ'], ['\\\\mu', 'μ'], ['\\\\pi', 'π'], ['\\\\sigma', 'σ'], ['\\\\omega', 'ω'],
  ];
  // 用字面量替换，不能把 LaTeX 反斜杠交给 RegExp 再解释；例如 `\\|` 会被正则当作“空匹配”。
  symbols.forEach(([from, to]) => { value = value.split(from).join(to); });

  // 兼容基础库对字符串反斜杠的差异：这里用明确的正则兜底替换所有常见命令。
  value = value
    // 预览只做可读降级：去除 Word/LaTex 的排版指令，但保留公式本身和编号。
    .replace(/\\(?:under|over)line\{((?:[^{}]|\{[^{}]*\})*)\}/g, '$1')
    .replace(/\\tag\{([^{}]*)\}/g, ' ($1)')
    .replace(/\\quad|\\qquad/g, ' ')
    .replace(/\\min\b/g, 'min').replace(/\\max\b/g, 'max')
    .replace(/\\bigcup/g, '⋃').replace(/\\bigcap/g, '⋂')
    .replace(/\\left|\\right|\\bigg|\\Bigg|\\big|\\Big/g, '')
    .replace(/\\langle/g, '⟨').replace(/\\rangle/g, '⟩')
    .replace(/\\sum/g, '∑').replace(/\\prod/g, '∏').replace(/\\int/g, '∫')
    .replace(/\\exists/g, '∃').replace(/\\forall/g, '∀').replace(/\\emptyset/g, '∅').replace(/\\infty/g, '∞')
    .replace(/\\subseteq/g, '⊆').replace(/\\supseteq/g, '⊇').replace(/\\subset/g, '⊂').replace(/\\supset/g, '⊃').replace(/\\notin/g, '∉').replace(/\\in/g, '∈')
    .replace(/\\leq|\\le/g, '≤').replace(/\\geq|\\ge/g, '≥').replace(/\\neq|\\ne/g, '≠')
    .replace(/\\cdots/g, '⋯').replace(/\\ldots|\\dots/g, '…').replace(/\\cap/g, '∩').replace(/\\cup/g, '∪')
    .replace(/\\Rightarrow/g, '⇒').replace(/\\rightarrow|\\to/g, '→').replace(/\\leftarrow/g, '←')
    .replace(/\\times/g, '×').replace(/\\cdot/g, '·')
    .replace(/\\Gamma/g, 'Γ').replace(/\\Delta/g, 'Δ').replace(/\\Theta/g, 'Θ').replace(/\\Psi/g, 'Ψ').replace(/\\Omega/g, 'Ω')
    .replace(/\\alpha/g, 'α').replace(/\\beta/g, 'β').replace(/\\gamma/g, 'γ').replace(/\\delta/g, 'δ').replace(/\\theta/g, 'θ').replace(/\\lambda/g, 'λ').replace(/\\mu/g, 'μ').replace(/\\pi/g, 'π').replace(/\\sigma/g, 'σ').replace(/\\omega/g, 'ω')
    // 帽子符号用组合字符保留，避免预览中出现 \hat y 这样的源码。
    .replace(/\\(?:widehat|hat)\s*\{?\s*([A-Za-z])\s*\}?/g, '$1̂')
    .replace(/\\,/g, ' ').replace(/\\\|/g, '|').replace(/\\\{/g, '{').replace(/\\\}/g, '}');

  // 清理 TeX 分组和空白控制符。下标/上标保留结构，以免复杂下标被错误篡改。
  value = value
    .replace(/\\[,!;:\s]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 小程序没有 TeX 排版引擎，至少把最常见的上下标变成可读的 Unicode 形式，
  // 不让预览页露出 H_{out}、x^2 这类源码。
  const lower = { a: 'ₐ', e: 'ₑ', h: 'ₕ', i: 'ᵢ', j: 'ⱼ', k: 'ₖ', l: 'ₗ', m: 'ₘ', n: 'ₙ', o: 'ₒ', p: 'ₚ', r: 'ᵣ', s: 'ₛ', t: 'ₜ', u: 'ᵤ', v: 'ᵥ', x: 'ₓ', '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄', '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉', '+': '₊', '-': '₋', '=': '₌' };
  const upper = { a: 'ᵃ', b: 'ᵇ', c: 'ᶜ', d: 'ᵈ', e: 'ᵉ', f: 'ᶠ', g: 'ᵍ', h: 'ʰ', i: 'ⁱ', j: 'ʲ', k: 'ᵏ', l: 'ˡ', m: 'ᵐ', n: 'ⁿ', o: 'ᵒ', p: 'ᵖ', r: 'ʳ', s: 'ˢ', t: 'ᵗ', u: 'ᵘ', v: 'ᵛ', w: 'ʷ', x: 'ˣ', y: 'ʸ', z: 'ᶻ', '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴', '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '+': '⁺', '-': '⁻', '=': '⁼' };
  const script = (chars, map, fallback) => chars.split('').map((char) => map[char.toLowerCase()] || fallback(char)).join('');
  const groupedScript = (chars, map, open, close) => {
    // 求和下标 i\in S_G 等是一个表达式，不能把其中的单个字母错误地逐字变成下标。
    if (!/^[A-Za-z0-9+\-=]+$/.test(chars)) return `${open}${chars}${close}`;
    if (!chars.split('').every((char) => char === char.toLowerCase() && map[char])) return `${open}${chars}${close}`;
    return script(chars, map, (char) => char);
  };
  return value
    .replace(/_\{([^{}]+)\}/g, (_, chars) => groupedScript(chars, lower, '₍', '₎'))
    .replace(/\^\{([^{}]+)\}/g, (_, chars) => groupedScript(chars, upper, '⁽', '⁾'))
    .replace(/_([A-Za-z0-9])/g, (_, char) => script(char, lower, (item) => `₍${item}₎`))
    .replace(/\^([A-Za-z0-9])/g, (_, char) => script(char, upper, (item) => `⁽${item}⁾`));
}

function renderInline(source) {
  const text = normalizeMathDelimiters(source);
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
  // 从 Markdown/聊天窗口复制时，公式分隔符可能被转义成 `\\(`、`\\[`；
  // 只还原分隔符，保留公式内部 `\\` 的数组换行语义。
  const input = normalizeMathDelimiters(String(source == null ? '' : source).replace(/\r\n?/g, '\n'));
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
      const rows = []; const first = i; const fence = trimmed.match(/^```\s*([^\s`]*)/); const language = fence && fence[1] ? fence[1].toLowerCase() : ''; i += 1;
      while (i < lines.length && !isCodeFence(lines[i])) { rows.push(lines[i]); i += 1; }
      const closed = i < lines.length; if (closed) i += 1;
      blocks.push(make('code', rows.join('\n'), start, blockEnd(Math.max(first, i - 1)), { closed, language }));
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
      // `\\(S=1\\), \\(K=3\\)` 是含两个行内公式的普通段落，不能因为行首是
      // `\\(` 就吞成一个未闭合的块公式。只有确实没有本行闭合符时才继续按多行公式读取。
      if (closeAt >= 0) {
        const rendered = renderInline(line);
        blocks.push(make('paragraph', line, start, blockEnd(i), { rendered })); i += 1; continue;
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

    // 大模型常省略 $...$ 分隔符；当一整行只包含数学字符/LaTeX 命令时，识别为公式块。
    const bareFormula = bareMathContent(line);
    if (bareFormula) {
      blocks.push(formulaBlock(bareFormula, start, blockEnd(i), true, 'bare-latex')); i += 1; continue;
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

    // Markdown 表格：表头下一行必须是 --- / :---: 形式，避免把普通的竖线文本误判为表格。
    if (isTableRow(line) && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
      const headers = tableCells(line); const separator = tableCells(lines[i + 1]); const first = i; i += 2;
      const rows = [];
      while (i < lines.length && isTableRow(lines[i]) && !isTableSeparator(lines[i])) {
        const cells = tableCells(lines[i]);
        if (cells.length < headers.length) while (cells.length < headers.length) cells.push('');
        rows.push(cells.slice(0, headers.length)); i += 1;
      }
      const alignments = separator.map((cell) => cell.startsWith(':') && cell.endsWith(':') ? 'center' : cell.endsWith(':') ? 'right' : 'left');
      const content = [headers].concat(rows);
      blocks.push(make('table', content, start, blockEnd(i - 1), { alignments, renderedRows: content.map((row) => row.map((cell) => renderInline(cell))) }));
      continue;
    }

    if (listLine(line)) {
      const rows = []; const renderedItems = []; const ordered = /^\s*\d+[.)]\s+/.test(line); const first = i;
      while (i < lines.length && listLine(lines[i])) {
        const item = wrapBareInlineVariables(lines[i].replace(/^\s*(?:[-*+]\s+|\d+[.)]\s+)/, '')); rows.push(item); renderedItems.push(renderInline(item)); i += 1;
      }
      blocks.push(make('list', rows, start, blockEnd(i - 1), { ordered, renderedItems })); continue;
    }

    const rows = [line.trim()]; const first = i; i += 1;
    while (i < lines.length) {
      const next = lines[i];
      if (!next.trim() || isCodeFence(next) || isMathFence(next) || isTableRow(next) || /^(\$\$|\\\[|\\\()/.test(next.trim()) || isStandaloneSingleDollar(next) || bareMathContent(next) || /^#{1,6}\s+/.test(next) || /^\s*>/.test(next) || listLine(next)) break;
      rows.push(next.trim()); i += 1;
    }
    const content = wrapBareInlineVariables(rows.join('\n')); blocks.push(make('paragraph', content, start, blockEnd(i - 1), { rendered: renderInline(content) }));
  }
  return blocks;
}

module.exports = { parseDocument, renderInline, normalizeLatex, formulaPreview, bareMathContent };
