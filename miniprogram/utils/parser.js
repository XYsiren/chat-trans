function makeId() { return `block_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`; }
function make(type, content, start, end, props) { return Object.assign({ id: makeId(), type, content, sourceRange: { start, end } }, props || {}); }
function listLine(line) { return /^\s*(?:[-*+]\s+|\d+[.)]\s+)/.test(line); }
function parseDocument(source) {
  const lines = String(source || '').replace(/\r\n?/g, '\n').split('\n'); const blocks = []; let i = 0; let pos = 0;
  while (i < lines.length) {
    const line = lines[i], start = pos; pos += line.length + 1;
    if (!line.trim()) { i++; continue; }
    if (/^\s*```/.test(line)) { const rows=[]; i++; while(i<lines.length&&!/^\s*```/.test(lines[i])) { rows.push(lines[i]); pos+=lines[i].length+1; i++; } if(i<lines.length) {pos+=lines[i].length+1;i++;} blocks.push(make('code',rows.join('\n'),start,pos)); continue; }
    const open=line.trim().match(/^(\$\$|\\\[)/); if(open) { const close=open[1]==='$$'?'$$':'\\]'; const rows=[]; let closed=false; i++; while(i<lines.length) { const row=lines[i]; pos+=row.length+1; if(row.trim()===close){closed=true;i++;break;} rows.push(row);i++; } blocks.push(make('formula',rows.join('\n').trim(),start,pos,{formulaFormat:'latex',closed}));continue; }
    const heading=line.match(/^(#{1,6})\s+(.+)$/); if(heading) {blocks.push(make('heading',heading[2],start,pos,{level:heading[1].length}));i++;continue;}
    if(/^\s*>/.test(line)) {const rows=[line.replace(/^\s*>\s?/, '')];i++;while(i<lines.length&&/^\s*>/.test(lines[i])){rows.push(lines[i].replace(/^\s*>\s?/,''));pos+=lines[i].length+1;i++;}blocks.push(make('quote',rows.join('\n'),start,pos));continue;}
    if(listLine(line)) {const rows=[];const ordered=/^\s*\d+[.)]\s+/.test(line);let row=line;while(row&&listLine(row)){rows.push(row.replace(/^\s*(?:[-*+]\s+|\d+[.)]\s+)/,''));i++;if(i>=lines.length)break;row=lines[i];if(listLine(row))pos+=row.length+1;}blocks.push(make('list',rows,start,pos,{ordered}));continue;}
    const rows=[line.trim()];i++;while(i<lines.length&&lines[i].trim()&&!/^\s*```/.test(lines[i])&&!/^(\$\$|\\\[)/.test(lines[i].trim())&&!/^#{1,6}\s+/.test(lines[i])&&!/^\s*>/.test(lines[i])&&!listLine(lines[i])){rows.push(lines[i].trim());pos+=lines[i].length+1;i++;}blocks.push(make('paragraph',rows.join('\n'),start,pos));
  } return blocks;
}
module.exports = { parseDocument };
