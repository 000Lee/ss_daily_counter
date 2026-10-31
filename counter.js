/* PT 시간표 카운터 — 핵심 로직 (DOM 의존 없음, JSZip + DOMParser만 사용)
 * 민감정보 없음. 업로드된 엑셀은 브라우저 메모리에서만 처리됨.
 */
(function (root) {
  'use strict';

  // ───────────────────────── 설정(사용자가 UI에서 수정 가능) ─────────────────────────
  const DEFAULTS = {
    // 1대다 접두 알파벳. 이 외의 접두 알파벳도 1대다로 처리하되 '기타 1대다'로 표시
    groupPrefixes: ['G', 'M'],
    // 확정 1대1 치료코드(접미). 그 외 접미/코드없음 → 1대1 '기타'
    oneCodes: ['N', 'G', 'M', 'Z', 'F'],
    // 장비·평가·주석으로 간주해 환자로 세지 않는 단어(부분 일치)
    ignoreWords: ['ESWT', '에리고', '족저압', '체평형', '신장', '로코멧', '모닝', 'sEMG', 'RG', '통증', '접수', '건수', '점심', '병동', '외진', '부도사유', 'GE',
      '조기퇴근', '퇴근', '연차', '반차', '휴가', '교육', '회의', '외출'],
    // 헤더(치료사명)에 괄호가 있으면 카운팅 제외 (예: 김효정B(통증), 신하린(8병동팀))
    excludeHeaderWithParen: true,
  };

  // ───────────────────────── XLSX 파싱 ─────────────────────────
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_DRAW = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  function parseXml(text) {
    return new DOMParser().parseFromString(text, 'application/xml');
  }
  function els(parent, tag, ns) {
    return Array.from(parent.getElementsByTagNameNS(ns || NS_MAIN, tag));
  }

  // 테마 색 배열: index 0=lt1,1=dk1,2=lt2,3=dk2,4..9=accent1..6
  function parseTheme(xml) {
    const out = [];
    if (!xml) return out;
    const doc = parseXml(xml);
    const scheme = els(doc, 'clrScheme', NS_DRAW)[0];
    if (!scheme) return out;
    const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    for (const name of order) {
      const el = els(scheme, name, NS_DRAW)[0];
      let rgb = null;
      if (el) {
        const sys = els(el, 'sysClr', NS_DRAW)[0];
        const srgb = els(el, 'srgbClr', NS_DRAW)[0];
        if (sys) rgb = sys.getAttribute('lastClr');
        else if (srgb) rgb = srgb.getAttribute('val');
      }
      out.push(rgb || (name.startsWith('lt') ? 'FFFFFF' : '000000'));
    }
    return out;
  }

  function hexToRgb(hex) {
    hex = hex.replace('#', '');
    if (hex.length === 8) hex = hex.slice(2); // ARGB
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16)];
  }
  function applyTint(rgb, tint) {
    if (!tint) return rgb;
    return rgb.map(c => {
      if (tint < 0) return Math.round(c * (1 + tint));
      return Math.round(c + (255 - c) * tint);
    });
  }
  function rgbToHex(rgb) {
    return '#' + rgb.map(c => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('');
  }

  // <color .../> 엘리먼트 → {hex, isGrey}
  function resolveColor(colorEl, theme) {
    if (!colorEl) return { hex: '#000000', isGrey: false, raw: 'auto' };
    let rgb = null, raw = '';
    if (colorEl.hasAttribute('rgb')) {
      rgb = hexToRgb(colorEl.getAttribute('rgb'));
      raw = 'rgb:' + colorEl.getAttribute('rgb');
    } else if (colorEl.hasAttribute('theme')) {
      const idx = parseInt(colorEl.getAttribute('theme'), 10);
      const tint = parseFloat(colorEl.getAttribute('tint') || '0');
      const base = theme[idx] || (idx === 0 ? 'FFFFFF' : '000000');
      rgb = applyTint(hexToRgb(base), tint);
      raw = `theme${idx}/${tint.toFixed(2)}`;
    } else if (colorEl.hasAttribute('indexed')) {
      const i = parseInt(colorEl.getAttribute('indexed'), 10);
      // 자주 쓰이는 indexed 색 일부
      const table = { 8: '000000', 9: 'FFFFFF', 10: 'FF0000', 12: '0000FF', 22: 'C0C0C0', 23: '808080', 55: '969696', 56: '333333', 63: '333333', 64: '000000' };
      rgb = hexToRgb(table[i] || '000000');
      raw = 'indexed:' + i;
    } else {
      return { hex: '#000000', isGrey: false, raw: 'auto' };
    }
    const [r, g, b] = rgb;
    const avg = (r + g + b) / 3;
    const spread = Math.max(r, g, b) - Math.min(r, g, b);
    const isGrey = spread <= 24 && avg >= 40 && avg <= 236; // 검정(<40)·흰색(>236)은 회색 아님
    return { hex: rgbToHex(rgb), isGrey, raw };
  }

  function parseFontEl(fontEl, theme) {
    const color = resolveColor(els(fontEl, 'color')[0], theme);
    const strike = els(fontEl, 'strike').length > 0;
    return { color: color.hex, isGrey: color.isGrey, strike, raw: color.raw };
  }

  function parseStyles(xml, theme) {
    const doc = parseXml(xml);
    const fonts = els(els(doc, 'fonts')[0] || doc, 'font').map(f => parseFontEl(f, theme));
    const cellXfs = els(doc, 'cellXfs')[0];
    const xfs = cellXfs ? els(cellXfs, 'xf').map(x => parseInt(x.getAttribute('fontId') || '0', 10)) : [];
    return { fonts, xfs };
  }

  // sharedStrings → [ [ {text, font|null} , ... ], ... ]  (font null = 셀 글꼴 상속)
  function parseSharedStrings(xml, theme) {
    if (!xml) return [];
    const doc = parseXml(xml);
    return els(doc, 'si').map(si => siToRuns(si, theme));
  }
  function siToRuns(si, theme) {
    const runs = [];
    for (const child of Array.from(si.children)) {
      const ln = child.localName;
      if (ln === 't') runs.push({ text: child.textContent, font: null });
      else if (ln === 'r') {
        const rPr = els(child, 'rPr')[0];
        const t = els(child, 't')[0];
        runs.push({ text: t ? t.textContent : '', font: rPr ? parseFontEl(rPr, theme) : null });
      }
    }
    return runs;
  }

  function colLetterToIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n; // 1-based
  }
  function indexToColLetter(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function parseSheet(xml, shared, styles) {
    const doc = parseXml(xml);
    const cells = {}; // 'B5' → {runs:[{text,font}], cellFont, text}
    let maxRow = 0, maxCol = 0;
    for (const c of els(doc, 'c')) {
      const ref = c.getAttribute('r');
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m) continue;
      const s = parseInt(c.getAttribute('s') || '0', 10);
      const fontId = styles.xfs[s] ?? 0;
      const cellFont = styles.fonts[fontId] || { color: '#000000', isGrey: false, strike: false };
      const t = c.getAttribute('t');
      const v = els(c, 'v')[0];
      let runs = null;
      if (t === 's' && v) runs = shared[parseInt(v.textContent, 10)] || [];
      else if (t === 'inlineStr') { const is = els(c, 'is')[0]; runs = is ? siToRuns(is, []) : []; }
      else if (v) runs = [{ text: v.textContent, font: null }];
      if (!runs || !runs.length) continue;
      // 글꼴 상속 처리
      runs = runs.map(r => ({ text: r.text, font: r.font || cellFont }));
      const text = runs.map(r => r.text).join('');
      if (!text.trim()) continue;
      cells[ref] = { ref, runs, cellFont, text };
      maxRow = Math.max(maxRow, parseInt(m[2], 10));
      maxCol = Math.max(maxCol, colLetterToIndex(m[1]));
    }
    return { cells, maxRow, maxCol };
  }

  async function readWorkbook(arrayBuffer, filename) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const get = async p => (zip.file(p) ? zip.file(p).async('string') : null);
    const theme = parseTheme(await get('xl/theme/theme1.xml'));
    const styles = parseStyles(await get('xl/styles.xml'), theme);
    const shared = parseSharedStrings(await get('xl/sharedStrings.xml'), theme);
    // 첫 시트
    const wbXml = parseXml(await get('xl/workbook.xml'));
    const relsXml = parseXml(await get('xl/_rels/workbook.xml.rels'));
    const firstSheet = els(wbXml, 'sheet')[0];
    let sheetPath = 'xl/worksheets/sheet1.xml';
    if (firstSheet) {
      const rid = firstSheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || firstSheet.getAttribute('r:id');
      const rel = Array.from(relsXml.getElementsByTagName('Relationship')).find(r => r.getAttribute('Id') === rid);
      if (rel) {
        let target = rel.getAttribute('Target');
        if (target.startsWith('/')) target = target.slice(1); else target = 'xl/' + target;
        sheetPath = target;
      }
    }
    const sheet = parseSheet(await get(sheetPath), shared, styles);
    return { filename, sheet };
  }

  // ───────────────────────── 시간표 구조 해석 ─────────────────────────
  const TIME_RE = /^\d{1,2}:\d{2}$/;
  function normTime(s) {
    const m = /^(\d{1,2}):(\d{2})$/.exec(String(s).trim());
    return m ? `${parseInt(m[1], 10)}:${m[2]}` : null;
  }

  function interpretSchedule(wb) {
    const { cells, maxRow, maxCol } = wb.sheet;
    const cellAt = (col, row) => cells[indexToColLetter(col) + row];

    // 날짜(A1), 팀(2행 또는 1~3행에서 'N팀' 찾기)
    const dateText = (cellAt(1, 1) || {}).text || '';
    let team = null;
    for (let r = 1; r <= 3 && team == null; r++) {
      for (let c = 1; c <= maxCol; c++) {
        const cell = cellAt(c, r);
        const m = cell && /(\d+)\s*팀/.exec(cell.text);
        if (m && !/FES/.test(cell.text)) { team = parseInt(m[1], 10); break; }
      }
    }
    // 헤더 행: 'FES\d+-\d' 가 있는 행
    let headerRow = 3;
    for (let r = 1; r <= 6; r++) {
      let found = false;
      for (let c = 1; c <= maxCol; c++) { const cell = cellAt(c, r); if (cell && /^FES\d+-\d/.test(cell.text.trim())) { found = true; break; } }
      if (found) { headerRow = r; break; }
    }
    // 열 분류
    const columns = []; // {col, letter, name, kind:'therapist'|'fes'|'ignore', fesKey, grey, excluded}
    let stop = false;
    for (let c = 2; c <= maxCol; c++) {
      const h = cellAt(c, headerRow);
      const letter = indexToColLetter(c);
      if (!h) { columns.push({ col: c, letter, name: '', kind: 'ignore', reason: '헤더 없음' }); continue; }
      const text = h.text.trim();
      const fes = /^FES(\d+)-(\d)/.exec(text);
      if (fes) { columns.push({ col: c, letter, name: text, kind: 'fes', fesKey: `${fes[1]}-${fes[2]}` }); continue; }
      if (/고정지원|검색|부도/.test(text)) { columns.push({ col: c, letter, name: text, kind: 'ignore', reason: text }); continue; }
      const grey = h.runs.every(r => r.font.isGrey);
      columns.push({ col: c, letter, name: text.replace(/\s+/g, ' '), kind: 'therapist', grey, hasParen: /[()（）]/.test(text) });
    }
    // 시간 행
    const rows = []; // {row, time, label}
    for (let r = headerRow + 1; r <= maxRow; r++) {
      const a = cellAt(1, r);
      if (!a) continue;
      const label = a.text.trim();
      const time = normTime(label);
      if (time) rows.push({ row: r, time, label });
      else if (/건수/.test(label)) break; // 건수 행 이후는 무시
      else rows.push({ row: r, time: null, label, skip: true }); // 점심 등
    }
    return { filename: wb.filename, dateText, team, headerRow, columns, rows, cells, cellAt };
  }

  // ───────────────────────── 셀 내용 → 라인 ─────────────────────────
  // 런(run)을 줄 단위로 쪼개 각 줄에 글꼴 속성 부여
  function cellToLines(cell) {
    if (!cell) return [];
    const lines = [];
    let cur = { text: '', fonts: [] };
    for (const run of cell.runs) {
      const parts = run.text.split('\n');
      parts.forEach((p, i) => {
        if (i > 0) { lines.push(cur); cur = { text: '', fonts: [] }; }
        cur.text += p;
        if (p.trim()) cur.fonts.push(run.font);
      });
    }
    lines.push(cur);
    return lines.map(l => {
      const f = l.fonts[0] || cell.cellFont;
      return { text: l.text.trim(), raw: l.text, color: f.color, isGrey: l.fonts.length ? l.fonts.every(x => x.isGrey) : f.isGrey, strike: l.fonts.length ? l.fonts.every(x => x.strike) : f.strike };
    }).filter(l => l.text.length > 0);
  }

  // ───────────────────────── 라인 해석 ─────────────────────────
  const NAME_RE = /^([A-Z]?)([가-힣]{2,4})\s?([A-Z]{0,2})$/; // 접두1 + 이름 + 접미0~2
  // 장비·주석 단어 판정: 영문 포함 단어는 부분일치, 한글 단어는 '단어로 시작하고 바로 뒤에 한글이 없을 때'(이름 오인 방지)
  function isIgnoredBy(core, w) {
    if (!w) return false;
    if (/[A-Za-z]/.test(w)) return core.toUpperCase().includes(w.toUpperCase());
    if (!core.startsWith(w)) return false;
    const next = core.charAt(w.length);
    return !/[가-힣]/.test(next);
  }
  function stripMarks(s) { return s.replace(/^[*\/\s]+/, '').replace(/[,，]+$/, '').trim(); }

  function classifyLine(line, cfg, mode) {
    // mode: 'therapist' | 'fes'(FES열: 이름만, 전부 환자)
    const text = line.text;
    const fesTokens = Array.from(text.matchAll(/FES\s*(\d+)\s*-\s*(\d)/g)).map(m => ({ team: parseInt(m[1], 10), slot: parseInt(m[2], 10), key: `${m[1]}-${m[2]}` }));
    if (fesTokens.length) return { type: 'fes', tokens: fesTokens, line };
    const core = stripMarks(text);
    if (!core) return { type: 'skip', line };
    if (cfg.ignoreWords.some(w => isIgnoredBy(core, w))) return { type: 'note', line };
    if (/^[\d\s:]+$/.test(core) || /^\(.*\)$/.test(core) || /^-?\d+팀$/.test(core) || /[0-9]/.test(core)) return { type: 'note', line };
    const m = NAME_RE.exec(core);
    if (!m) return { type: 'note', line, odd: true };
    const [, prefix, name, suffix] = m;
    const hadSlash = /^\s*[*\/]/.test(text);
    if (mode === 'fes') {
      // FES열: 접두/접미 무시, 이름만 환자. 동명이인 표기는 접미 1글자
      return { type: 'patient', kind: 'group', name, dup: suffix.slice(0, 1) || '', code: 'FES', line, hadSlash };
    }
    if (prefix) {
      const known = cfg.groupPrefixes.includes(prefix);
      return { type: 'patient', kind: 'group', name, dup: suffix.slice(0, 1) || '', code: prefix, known, line, hadSlash };
    }
    let dup = '', code = '';
    if (suffix.length === 2) { dup = suffix[0]; code = suffix[1]; }
    else if (suffix.length === 1) { code = suffix; }
    const known = code ? cfg.oneCodes.includes(code) : false;
    return { type: 'patient', kind: 'one', name, dup, code, known, isFesOne: code === 'F', line, hadSlash, shortName: name.length === 2 };
  }

  // ───────────────────────── 카운팅 ─────────────────────────
  function weightOf(line) {
    if (line.isGrey) return 0;      // 부도
    if (line.strike) return 0.5;    // 반부도(검정 취소선)
    return 1;
  }

  /**
   * schedules: interpretSchedule 결과 배열(여러 팀)
   * targetTeam: 카운팅할 팀 번호
   * cfg: 설정
   */
  function analyze(schedules, targetTeam, cfg) {
    cfg = Object.assign({}, DEFAULTS, cfg || {});
    const target = schedules.find(s => s.team === targetTeam);
    if (!target) throw new Error(`${targetTeam}팀 시간표가 없습니다`);
    const byTeam = {}; schedules.forEach(s => { byTeam[s.team] = s; });

    // FES 조회: 팀/슬롯/시간 → 환자 항목들
    function fesPatients(team, slot, time) {
      const sch = byTeam[team];
      if (!sch) return { missing: true, items: [] };
      const col = sch.columns.find(c => c.kind === 'fes' && c.fesKey === `${team}-${slot}`);
      const row = sch.rows.find(r => r.time === time);
      if (!col || !row) return { missing: true, items: [], detail: !col ? `FES${team}-${slot} 열 없음` : `${time} 행 없음` };
      const cell = sch.cellAt(col.col, row.row);
      const items = [];
      for (const ln of cellToLines(cell)) {
        const c = classifyLine(ln, cfg, 'fes');
        if (c.type === 'patient') items.push({ ...c, weight: weightOf(ln), src: `${team}팀 ${col.letter}${row.row}` });
      }
      return { missing: false, items, ref: `${team}팀 ${col.letter}${row.row}`, empty: !cell };
    }

    const therapists = [];
    const warnings = [];
    for (const col of target.columns) {
      if (col.kind !== 'therapist') continue;
      const excluded = col.grey || (cfg.excludeHeaderWithParen && col.hasParen) || (cfg.excludeColumns || []).includes(col.letter);
      const th = { col, name: col.name, excluded, excludeReason: col.grey ? '연차(회색 헤더)' : (col.hasParen ? '헤더 괄호' : (excluded ? '수동 제외' : '')), one: 0, group: 0, oneOther: 0, oneFes: 0, groupOther: 0, cells: [] };
      for (const row of target.rows) {
        const cell = target.cellAt(col.col, row.row);
        const entry = { row, ref: `${col.letter}${row.row}`, lines: [], one: 0, group: 0, skipped: !!row.skip };
        if (cell && !row.skip) {
          for (const ln of cellToLines(cell)) {
            const c = classifyLine(ln, cfg, 'therapist');
            const w = weightOf(ln);
            const out = { text: ln.text, color: ln.color, isGrey: ln.isGrey, strike: ln.strike, type: c.type, weight: w, tag: '' };
            if (c.type === 'patient') {
              out.name = c.name; out.code = c.code; out.dup = c.dup; out.kind = c.kind; out.known = c.known;
              if (c.kind === 'one') {
                out.tag = w === 0 ? '부도' : (w === 0.5 ? '1:1 ½' : '1:1');
                if (c.isFesOne) out.tag += ' F';
                if (!c.known && w > 0) { out.tag += ' 기타'; out.other = true; }
                if (c.shortName && w > 0) { out.odd = true; }
                entry.one += w;
                if (!excluded) { th.one += w; if (out.other) th.oneOther += w; if (c.isFesOne) th.oneFes += w; }
              } else {
                out.tag = w === 0 ? '부도' : (w === 0.5 ? '1:다 ½' : '1:다');
                if (!c.known && w > 0) { out.tag += ' 기타'; out.other = true; }
                entry.group += w;
                if (!excluded) { th.group += w; if (out.other) th.groupOther += w; }
              }
            } else if (c.type === 'fes') {
              out.fes = [];
              if (w === 0) { out.tag = 'FES 부도'; }
              else {
                for (const tk of c.tokens) {
                  const res = fesPatients(tk.team, tk.slot, row.time);
                  const f = { key: `FES${tk.key}`, missing: res.missing, detail: res.detail, ref: res.ref, items: res.items.map(it => ({ name: it.name, dup: it.dup, weight: it.weight * w, isGrey: it.line.isGrey, strike: it.line.strike })) };
                  const sum = f.items.reduce((a, b) => a + b.weight, 0);
                  f.sum = sum;
                  entry.group += sum;
                  if (!excluded) th.group += sum;
                  if (res.missing) warnings.push(`${col.name} ${entry.ref}: ${f.key} → ${res.detail || `${tk.team}팀 파일 없음`}`);
                  out.fes.push(f);
                }
                out.tag = '1:다 ' + out.fes.map(f => f.sum).reduce((a, b) => a + b, 0);
              }
            } else if (c.type === 'note') {
              out.tag = '';
            }
            entry.lines.push(out);
          }
        }
        th.cells.push(entry);
      }
      therapists.push(th);
    }
    return { target, therapists, warnings, cfg };
  }

  root.PTCounter = { DEFAULTS, readWorkbook, interpretSchedule, cellToLines, classifyLine, analyze, normTime };
})(typeof window !== 'undefined' ? window : globalThis);
