// ============================================================
// Word 課表匯出 (班級課表 + 教師課表)
// 模板：class-official-template.docx / class-teacher-official-template.docx
// 依賴：PizZip、FileSaver.js
// 鐵則：只做字串替換佔位符，不改字型／粗體／段落／列高
// ============================================================

function parseDocxZip(buf) {
  const ZipConstructor = (typeof PizZip !== 'undefined') ? PizZip : ((typeof window !== 'undefined' && window.PizZip) ? window.PizZip : (typeof JSZip !== 'undefined' ? JSZip : null));
  if (!ZipConstructor) throw new Error('找不到 Word 解壓套件 (PizZip)，請確認網路連線正常後重新整理頁面');
  return new ZipConstructor(buf);
}

let _tplCache = null;
let _teacherTplCache = null;
let _roomTplCache = null;
let _wordCurrentTab = 'class'; // 'class' | 'teacher' | 'room' | 'patrol'

const BUDING_SUBJECTS = new Set([
  '國文', '英語', '英文', '本土語',
  '數學',
  '歷史', '地理', '公民', '公民與社會',
  '生物', '理化', '自然', '地球科學',
  '音樂', '視覺藝術', '表演藝術',
  '家政', '童軍', '輔導',
  '資訊科技', '生活科技',
  '健康教育', '體育'
]);

const FLEX_SUBJECT_ORDER = [
  '走讀建成生活圈', '全球議題', '文旅享繪', '活力建成',
  '建成公民行動家', '文化種籽在建成', '全球素養',
  '英悅讀樂樂', '閱思溝通建成人', '藝統摺學', '全民國防'
];

const SUBJECT_ALIASES = {
  '公民與社會': '公民', '公民': '公民',
  '英文': '英語', '週會': '班週會',
  '理化': '生物', '自然': '生物'
};

const FIXED_SLOT_NAMES = [
  '國文', '生物', '班週會', '英語', '家政', '本土語', '童軍',
  '數學', '輔導', '資訊科技', '音樂', '生活科技', '視覺藝術',
  '歷史', '表演藝術', '地理', '健康教育', '公民', '體育'
];

const WORD_EARLY_PERIOD = 0;
const WORD_LUNCH_PERIOD = 45;

function wordRowText(xml) {
  return String(xml || '')
    .replace(/<w:tab\s*\/?\s*>/g, '\t')
    .replace(/<w:br\s*\/?\s*>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function wordRows(xml) {
  return [...String(xml || '').matchAll(/<w:tr(?:\s[^>]*)?>[\s\S]*?<\/w:tr>/g)].map(match => match[0]);
}

function wordCells(rowXml) {
  return [...String(rowXml || '').matchAll(/<w:tc(?:\s[^>]*)?>[\s\S]*?<\/w:tc>/g)].map(match => match[0]);
}

function wordReplacePlaceholders(cellXml, replacements) {
  let out = String(cellXml || '');
  Object.entries(replacements || {}).forEach(([from, to]) => {
    out = out.split('{' + from + '}').join('{' + to + '}');
  });
  return out;
}

function wordBlankCell(cellXml) {
  return String(cellXml || '').replace(/<w:t(\s[^>]*)?>[\s\S]*?<\/w:t>/g, '<w:t$1></w:t>');
}

function wordRemoveBold(cellXml) {
  return String(cellXml || '')
    .replace(/<w:b(?:\s[^>]*)?\s*\/?>(?:<\/w:b>)?/g, '')
    .replace(/<w:bCs(?:\s[^>]*)?\s*\/?>(?:<\/w:bCs>)?/g, '');
}

function wordSetVerticalMerge(cellXml, mode) {
  let out = String(cellXml || '');
  const merge = mode === 'restart' ? '<w:vMerge w:val="restart"/>' : (mode === 'continue' ? '<w:vMerge/>' : '');
  out = out.replace(/<w:vMerge(?:\s[^>]*)?\s*\/?>(?:<\/w:vMerge>)?/g, '');
  if (!merge) return out;
  const tcPr = out.match(/<w:tcPr[\s\S]*?<\/w:tcPr>/);
  if (tcPr) return out.replace(tcPr[0], tcPr[0].replace('</w:tcPr>', merge + '</w:tcPr>'));
  return out.replace(/(<w:tc(?:\s[^>]*)?>)/, '$1<w:tcPr>' + merge + '</w:tcPr>');
}

function wordRebuildRow(rowXml, cells) {
  const open = String(rowXml || '').match(/^<w:tr[^>]*>/)?.[0] || '<w:tr>';
  const trPr = String(rowXml || '').match(/<w:trPr[\s\S]*?<\/w:trPr>/)?.[0] || '';
  return open + trPr + cells.join('') + '</w:tr>';
}

function wordSpecialValue(dict, day, period, mode, suffix) {
  const key = mode === 'class'
    ? 'd' + day + 'p' + period
    : 'd' + day + 'p' + period + '_' + suffix;
  return String(dict[key] || '').trim();
}

function expandWordSpecialRows(pageXml, dict, mode) {
  const rowList = wordRows(pageXml);
  const sourceRow = rowList.find(row => mode === 'class' ? /\{d1p1\}/.test(row) : /\{d1p1_s\}/.test(row));
  if (!sourceRow) return pageXml;
  const sourceCells = wordCells(sourceRow);
  const dayStart = mode === 'class' ? 2 : 3;
  if (sourceCells.length < dayStart + 5) return pageXml;

  let output = pageXml;
  const buildRow = (rowXml, period, topRow, mergeDays) => {
    const targetCells = wordCells(rowXml);
    if (targetCells.length < 2) return rowXml;
    const cells = [targetCells[0]];
    for (let day = 1; day <= 5; day++) {
      const sourceCell = sourceCells[dayStart + day - 1];
      const replacements = mode === 'class'
        ? { ['d' + day + 'p1']: 'd' + day + 'p' + period }
        : {
            ['d' + day + 'p1_s']: 'd' + day + 'p' + period + '_s',
            ['d' + day + 'p1_c']: 'd' + day + 'p' + period + '_c'
          };
      let cell = wordReplacePlaceholders(sourceCell, replacements);
      cell = wordRemoveBold(cell);
      const hasValue = wordSpecialValue(dict, day, period, mode, 's') || wordSpecialValue(dict, day, period, mode, '');
      if (!topRow) cell = wordBlankCell(cell);
      if (mergeDays && hasValue) cell = wordSetVerticalMerge(cell, topRow ? 'restart' : 'continue');
      cells.push(cell);
    }
    return wordRebuildRow(rowXml, cells);
  };

  const earlyTop = rowList.find(row => wordRowText(row).includes('07:40'));
  const earlyBottom = rowList.find(row => wordRowText(row).includes('08:15') && !wordRowText(row).includes('07:40'));
  const earlyValues = Array.from({ length: 5 }, (_, index) => wordSpecialValue(dict, index + 1, WORD_EARLY_PERIOD, mode, 's') || wordSpecialValue(dict, index + 1, WORD_EARLY_PERIOD, mode, ''));
  if (earlyTop && earlyBottom && earlyValues.some(Boolean)) {
    output = output.replace(earlyTop, buildRow(earlyTop, WORD_EARLY_PERIOD, true, true));
    output = output.replace(earlyBottom, buildRow(earlyBottom, WORD_EARLY_PERIOD, false, true));
  }

  const lunchRow = wordRows(output).find(row => wordRowText(row).includes('12:35') && wordRowText(row).includes('午休'));
  const lunchValues = Array.from({ length: 5 }, (_, index) => wordSpecialValue(dict, index + 1, WORD_LUNCH_PERIOD, mode, 's') || wordSpecialValue(dict, index + 1, WORD_LUNCH_PERIOD, mode, ''));
  if (lunchRow && lunchValues.some(Boolean)) {
    output = output.replace(lunchRow, buildRow(lunchRow, WORD_LUNCH_PERIOD, true, false));
  }
  return output;
}

function switchWordTab(tab) {
  _wordCurrentTab = tab;
  const classTabBtn = document.getElementById('word-tab-class');
  const teacherTabBtn = document.getElementById('word-tab-teacher');
  const roomTabBtn = document.getElementById('word-tab-room');
  const patrolTabBtn = document.getElementById('word-tab-patrol');
  const classPanel = document.getElementById('word-class-panel');
  const teacherPanel = document.getElementById('word-teacher-panel');
  const roomPanel = document.getElementById('word-room-panel');
  const patrolPanel = document.getElementById('word-patrol-panel');
  const selectControls = document.getElementById('word-select-controls');
  const desc = document.getElementById('word-desc');

  if (tab === 'class') {
    if (classTabBtn) classTabBtn.className = 'btn btn-sm btn-primary';
    if (teacherTabBtn) teacherTabBtn.className = 'btn btn-sm btn-ghost';
    if (roomTabBtn) roomTabBtn.className = 'btn btn-sm btn-ghost';
    if (patrolTabBtn) patrolTabBtn.className = 'btn btn-sm btn-ghost';
    if (classPanel) classPanel.style.display = 'block';
    if (teacherPanel) teacherPanel.style.display = 'none';
    if (roomPanel) roomPanel.style.display = 'none';
    if (patrolPanel) patrolPanel.style.display = 'none';
    if (selectControls) selectControls.style.display = 'flex';
    if (desc) desc.textContent = '勾選要匯出的班級，套用官方 Word 範本產出一份 .docx（每班一頁）。';
  } else if (tab === 'teacher') {
    if (classTabBtn) classTabBtn.className = 'btn btn-sm btn-ghost';
    if (teacherTabBtn) teacherTabBtn.className = 'btn btn-sm btn-primary';
    if (roomTabBtn) roomTabBtn.className = 'btn btn-sm btn-ghost';
    if (patrolTabBtn) patrolTabBtn.className = 'btn btn-sm btn-ghost';
    if (classPanel) classPanel.style.display = 'none';
    if (teacherPanel) teacherPanel.style.display = 'block';
    if (roomPanel) roomPanel.style.display = 'none';
    if (patrolPanel) patrolPanel.style.display = 'none';
    if (selectControls) selectControls.style.display = 'flex';
    if (desc) desc.textContent = '勾選要匯出的教師，套用官方 Word 範本產出一份 .docx（每人一頁，含配課總表）。第八節支援單雙週拆欄。';
  } else if (tab === 'room') {
    if (classTabBtn) classTabBtn.className = 'btn btn-sm btn-ghost';
    if (teacherTabBtn) teacherTabBtn.className = 'btn btn-sm btn-ghost';
    if (roomTabBtn) roomTabBtn.className = 'btn btn-sm btn-primary';
    if (patrolTabBtn) patrolTabBtn.className = 'btn btn-sm btn-ghost';
    if (classPanel) classPanel.style.display = 'none';
    if (teacherPanel) teacherPanel.style.display = 'none';
    if (roomPanel) roomPanel.style.display = 'block';
    if (patrolPanel) patrolPanel.style.display = 'none';
    if (selectControls) selectControls.style.display = 'flex';
    if (desc) desc.textContent = '勾選要匯出的專科教室，套用官方 Word 範本 room-official-template.docx 產出一份 .docx（每間教室一頁，含使用總表）。';
  } else {
    if (classTabBtn) classTabBtn.className = 'btn btn-sm btn-ghost';
    if (teacherTabBtn) teacherTabBtn.className = 'btn btn-sm btn-ghost';
    if (roomTabBtn) roomTabBtn.className = 'btn btn-sm btn-ghost';
    if (patrolTabBtn) patrolTabBtn.className = 'btn btn-sm btn-primary';
    if (classPanel) classPanel.style.display = 'none';
    if (teacherPanel) teacherPanel.style.display = 'none';
    if (roomPanel) roomPanel.style.display = 'none';
    if (patrolPanel) patrolPanel.style.display = 'block';
    if (selectControls) selectControls.style.display = 'none';
    if (desc) desc.textContent = '匯出全部巡堂時段，套用 Word 表格產出一份巡堂表。';
  }
  updateWordCount();
}

function openWordExportModal() {
  const classList = document.getElementById('word-class-list');
  const teacherList = document.getElementById('word-teacher-list');
  const roomList = document.getElementById('word-room-list');
  const progress = document.getElementById('word-export-progress');
  const btn = document.getElementById('word-export-btn');

  if (progress) progress.style.display = 'none';
  if (btn) {
    btn.disabled = false;
    btn.textContent = '📥 開始匯出';
  }

  // 1. 渲染班級清單
  const realClasses = state.classes || [];

  if (classList) {
    classList.innerHTML = '';
    realClasses.forEach(cls => {
      const code = cls['班級代碼'];
      const name = cls['班級名稱'];
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;padding:5px 8px;border:1px solid var(--border);border-radius:6px;';
      const virtualMark = String(cls['是否虛擬班']).toUpperCase() === 'TRUE' ? '（抽離／虛擬）' : '';
      label.innerHTML = `<input type="checkbox" class="word-cls-chk" value="${code}" checked onchange="updateWordCount()"> ${name}${virtualMark}`;
      classList.appendChild(label);
    });
  }

  // 2. 渲染教師清單
  if (teacherList) {
    teacherList.innerHTML = '';
    (state.teachers || []).forEach(t => {
      const code = t['姓名'] || t['教師姓名'];
      const name = code;
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;padding:5px 8px;border:1px solid var(--border);border-radius:6px;';
      label.innerHTML = `<input type="checkbox" class="word-t-chk" value="${code}" checked onchange="updateWordCount()"> ${name}`;
      teacherList.appendChild(label);
    });
  }

  // 3. 渲染教室清單
  if (roomList) {
    roomList.innerHTML = '';
    (state.rooms || []).forEach(r => {
      const code = String(r['教室代碼'] || '');
      const name = r['教室名稱'] || code;
      const label = document.createElement('label');
      label.style.cssText = 'display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;padding:5px 8px;border:1px solid var(--border);border-radius:6px;';
      label.innerHTML = `<input type="checkbox" class="word-r-chk" value="${code}" checked onchange="updateWordCount()"> ${name}`;
      roomList.appendChild(label);
    });
  }

  const patrolList = document.getElementById('word-patrol-list');
  const patrolSummary = document.getElementById('word-patrol-summary');
  const patrolRows = (state.schedule || [])
    .filter(isPatrolScheduleEntry)
    .sort((a, b) => parseInt(a['星期'], 10) - parseInt(b['星期'], 10) ||
      parseInt(a['節次'], 10) - parseInt(b['節次'], 10) ||
      String(a['教師姓名'] || '').localeCompare(String(b['教師姓名'] || ''), 'zh-Hant'));
  if (patrolSummary) patrolSummary.textContent = patrolRows.length ? '共 ' + patrolRows.length + ' 節，匯出時會全部列入巡堂表。' : '目前尚未建立巡堂時段。';
  if (patrolList) {
    const dayNames = ['', '週一', '週二', '週三', '週四', '週五'];
    patrolList.innerHTML = patrolRows.length
      ? '<table style="width:100%;border-collapse:collapse;font-size:13px;"><thead><tr><th style="text-align:left;padding:6px;border-bottom:1px solid var(--border);">星期</th><th style="text-align:left;padding:6px;border-bottom:1px solid var(--border);">節次</th><th style="text-align:left;padding:6px;border-bottom:1px solid var(--border);">巡堂教師</th><th style="text-align:left;padding:6px;border-bottom:1px solid var(--border);">職務</th></tr></thead><tbody>' +
        patrolRows.map(row => {
          const code = String(row['教師姓名'] || '').trim();
          const teacher = idx.teacherByCode[code] || {};
          const title = teacher['職務'] || teacher['職稱'] || '';
          return '<tr><td style="padding:6px;border-bottom:1px solid var(--border);">' + esc(dayNames[parseInt(row['星期'], 10)] || row['星期']) + '</td>' +
            '<td style="padding:6px;border-bottom:1px solid var(--border);">第' + esc(row['節次']) + '節</td>' +
            '<td style="padding:6px;border-bottom:1px solid var(--border);font-weight:bold;">' + esc(teacherName(code)) + '</td>' +
            '<td style="padding:6px;border-bottom:1px solid var(--border);">' + esc(title) + '</td></tr>';
        }).join('') + '</tbody></table>'
      : '<div class="text-muted" style="padding:18px 6px;text-align:center;">目前沒有巡堂時段</div>';
  }

  switchWordTab(_wordCurrentTab || 'class');
  document.getElementById('wordExportModal').classList.add('show');
}

function closeWordExportModal() {
  document.getElementById('wordExportModal').classList.remove('show');
}

function toggleWordSelectAll(checked) {
  if (_wordCurrentTab === 'class') {
    document.querySelectorAll('.word-cls-chk').forEach(chk => { chk.checked = checked; });
  } else if (_wordCurrentTab === 'teacher') {
    document.querySelectorAll('.word-t-chk').forEach(chk => { chk.checked = checked; });
  } else if (_wordCurrentTab === 'room') {
    document.querySelectorAll('.word-r-chk').forEach(chk => { chk.checked = checked; });
  }
  updateWordCount();
}

function updateWordCount() {
  const tab = _wordCurrentTab || 'class';
  if (tab === 'patrol') {
    const patrolCount = (state.schedule || []).filter(isPatrolScheduleEntry).length;
    const cntEl = document.getElementById('word-selected-count');
    if (cntEl) cntEl.textContent = '全部 ' + patrolCount + ' 節巡堂';
    return;
  }
  const selector = tab === 'class' ? '.word-cls-chk' : (tab === 'teacher' ? '.word-t-chk' : '.word-r-chk');
  const selected = document.querySelectorAll(selector + ':checked').length;
  const total = document.querySelectorAll(selector).length;
  const unit = tab === 'class' ? '班' : (tab === 'teacher' ? '位教師' : '間教室');
  
  const cntEl = document.getElementById('word-selected-count');
  if (cntEl) cntEl.textContent = `已選 ${selected} / ${total} ${unit}`;

  const allChk = document.getElementById('word-select-all');
  if (allChk) {
    allChk.checked = selected === total && total > 0;
    allChk.indeterminate = selected > 0 && selected < total;
  }
}

function escXml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function teacherName(code) {
  if (!code) return '';
  const t = idx.teacherByCode[code];
  return t ? (t['教師姓名'] || t['姓名'] || code) : String(code);
}

function resolveHomeTeacher(classCode, classInfo) {
  let code = String((classInfo && classInfo['導師代碼']) || '').trim();
  if (!code) {
    const ht = (state.teachers || []).find(t => {
      const hr = (typeof getTeacherHomeroom === 'function')
        ? getTeacherHomeroom(t)
        : String(t['導師班級'] || '').trim();
      return hr && hr !== 'TRUE' && String(hr) === String(classCode);
    });
    if (ht) code = String(ht['教師姓名'] || '');
  }
  return teacherName(code) || code;
}

function normalizeSubjectKey(subj) {
  let s = String(subj || '').trim();
  s = s.replace(/（/g, '(').replace(/）/g, ')').replace(/\s+/g, '');
  if (SUBJECT_ALIASES[s]) return SUBJECT_ALIASES[s];
  return s;
}

function stripSubjectNoise(subj) {
  return String(subj || '')
    .trim()
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/\s+/g, '')
    .replace(/\(.*?\)/g, '');
}

function isBudingSubject(subj) {
  const raw = String(subj || '').trim();
  const key = normalizeSubjectKey(raw);
  const bare = stripSubjectNoise(raw);
  if (BUDING_SUBJECTS.has(raw) || BUDING_SUBJECTS.has(key) || BUDING_SUBJECTS.has(bare)) return true;
  if (/輔$|^\S+輔$|單$|雙$/.test(bare)) return true;
  return false;
}

function isBanZhouHui(subj) {
  const bare = stripSubjectNoise(subj);
  return bare === '班週會' || bare === '週會' || normalizeSubjectKey(subj) === '班週會';
}

function isFlexSubject(subj) {
  if (isBanZhouHui(subj)) return false;
  if (isBudingSubject(subj)) return false;
  return true;
}

function findCourseFuzzy(courses, want) {
  const w = stripSubjectNoise(want);
  if (courses[want]) return courses[want];
  const nk = normalizeSubjectKey(want);
  if (courses[nk]) return courses[nk];
  for (const k of Object.keys(courses)) {
    const bare = stripSubjectNoise(courses[k].subject || k);
    if (bare === w || bare.startsWith(w) || w.startsWith(bare)) return courses[k];
    if (stripSubjectNoise(k) === w) return courses[k];
  }
  return null;
}

function collectClassCourses(classCode) {
  const map = {};

  function nameOf(code) {
    const t = idx.teacherByCode[String(code || '')];
    return t ? (t['姓名'] || String(code || '')) : String(code || '');
  }
  // 每個 key 維護 teacherByCode → 顯示字（含標籤）；以代碼為準，最後寫入覆蓋
  function ensureKey(raw, sub) {
    const key = normalizeSubjectKey(raw);
    if (!map[key]) {
      map[key] = { subject: raw, teachers: [], periods: 0, _byCode: {} };
    }
    if (sub.length > map[key].subject.length) map[key].subject = sub;
    return key;
  }
  function upsert(subj, teacherCode, tag, periods) {
    const raw = String(subj || '').trim();
    if (!raw) return;
    const key = ensureKey(raw, raw);
    const c = String(teacherCode || '').trim();
    if (c) {
      const name = nameOf(c);
      if (name) map[key]._byCode[c] = String(tag || '').trim() ? name + '（' + String(tag).trim() + '）' : name;
    }
    if (periods != null && periods !== '') {
      const n = parseInt(periods, 10);
      if (!isNaN(n) && n > 0) map[key].periods = Math.max(map[key].periods, n);
    }
  }

  (state.assignments || []).forEach(a => {
    if (String(a['班級代碼']) !== String(classCode)) return;
    const sub = a['科目代碼'] || '';
    const tc = a['教師姓名'] || '';
    const info = idx.subjectByCode[sub];
    const weekly = a['每週節數'] || (info ? info['每週節數'] : '') || '';
    upsert(sub, tc, '', weekly);
  });

  const bySub = {};
  (state.schedule || []).forEach(s => {
    if (String(s['班級代碼']) !== String(classCode)) return;
    const sub = String(s['科目代碼'] || '').trim();
    if (!sub) return;
    const tc = String(s['教師姓名'] || '').trim();
    // 多教師：收集「教師代碼」欄解析出的每位教師（含標籤），供配課總表分欄呈現
    const tList = getCellTeacherList(s);
    if (tList.length > 0) {
      const key = ensureKey(sub, sub);
      tList.forEach(t => {
        const c = String(t['教師姓名'] || '').trim();
        if (!c) return;
        const name = nameOf(c);
        if (!name) return;
        const labelled = (String(t['標籤'] || '').trim())
          ? name + '（' + String(t['標籤']).trim() + '）'
          : name;
        map[key]._byCode[c] = labelled;
      });
    } else {
      upsert(sub, tc, '', 0);
    }
    const k = normalizeSubjectKey(sub);
    bySub[k] = (bySub[k] || 0) + 1;
  });
  // 依代碼去重：物化 teachers（保留科目出現在順序）
  Object.keys(map).forEach(k => {
    map[k].teachers = Object.keys(map[k]._byCode).map(c => map[k]._byCode[c]);
    delete map[k]._byCode;
  });
  Object.keys(bySub).forEach(k => {
    if (map[k] && (!map[k].periods || map[k].periods < bySub[k])) map[k].periods = bySub[k];
  });

  return map;
}

function isWordOvertimeCell(cell) {
  return !!cell && String(cell['課堂屬性'] || '').trim() === '超鐘點';
}

function teacherWordSubject(cell) {
  if (typeof isPatrolScheduleEntry === 'function' && isPatrolScheduleEntry(cell)) return '巡堂';
  const subject = String(cell && cell['科目代碼'] || '').trim();
  return subject + (isWordOvertimeCell(cell) ? '（超）' : '');
}

function teacherWordSpecialSubject(cells) {
  const grouped = {};
  (cells || []).forEach(cell => {
    const subject = teacherWordSubject(cell);
    if (!subject) return;
    if (!grouped[subject]) grouped[subject] = new Set();
    const classCode = String(cell && cell['班級代碼'] || '').trim();
    if (classCode) grouped[subject].add(classCode);
  });
  return Object.keys(grouped).map(subject => {
    const classes = [...grouped[subject]];
    return classes.length ? subject + '（' + classes.join('、') + '）' : subject;
  }).join('／');
}

function slotSubject(classCode, day, period) {
  if (period === 8) {
    // 第八節回傳合併顯示（向後相容）
    const p8 = idx.schedByClassSlotP8[`${classCode}|${day}|8`] || {};
    const s = p8['單週'] ? String(p8['單週']['科目代碼'] || '') + '(單)' : '';
    const d = p8['雙週'] ? String(p8['雙週']['科目代碼'] || '') + '(雙)' : '';
    const parts = [s, d].filter(Boolean);
    return parts.join(' / ');
  }
  const cell = idx.schedByClassSlot[`${classCode}|${day}|${period}`];
  return cell && cell['科目代碼'] ? String(cell['科目代碼']) : '';
}

function slotSubjectP8(classCode, day, weekType) {
  const p8 = idx.schedByClassSlotP8[`${classCode}|${day}|8`] || {};
  const cell = p8[weekType];
  return cell && cell['科目代碼'] ? String(cell['科目代碼']) : '';
}

function joinSplitPlaceholders(xml) {
  const pairs = [
    ['{生活科', '技節}'],
    ['{視覺藝', '術節}'],
    ['{', '姓名}']
  ];
  let out = xml;
  pairs.forEach(([pre, suf]) => {
    const re = new RegExp(escRegex(pre) + '(</w:t></w:r>\\s*<w:r[ >][\\s\\S]*?<w:t[^>]*>)' + escRegex(suf), 'g');
    out = out.replace(re, pre + suf);
  });
  return out;
}

function escRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function fillPlaceholders(xml, dict) {
  let out = xml;
  Object.keys(dict)
    .sort((a, b) => b.length - a.length)
    .forEach(k => {
      let val = escXml(dict[k] == null ? '' : dict[k]);
      val = val.replace(/\n/g, '<w:br/>');
      out = out.split('{' + k + '}').join(val);
    });
  out = out.replace(/\{[a-zA-Z0-9\u4e00-\u9fff_]+\}/g, '');
  return out;
}

// 為課表格子注入底色：findMap 為 { 佔位符鍵: 底色HEX }；只處理有底色的格子
function injectCellFills(xml, fillMap) {
  if (!fillMap || !Object.keys(fillMap).length) return xml;
  // 找出 XML 中所有 w:tc 區段，只針對含「填色鍵」佔位符者加 w:shd
  return xml.replace(/(<w:tc\b[^>]*>)([\s\S]*?)(<\/w:tc>)/g, (whole, open, inner, close) => {
    const keys = inner.match(/\{([a-zA-Z0-9_\u4e00-\u9fff]+)\}/g);
    if (!keys) return whole;
    let fill = '';
    keys.forEach(k => {
      const key = k.slice(1, -1);
      const f = (fillMap[key] || '').trim();
      if (f) fill = f; // 終以 find 順序，取最後一個有值者
    });
    if (!fill) return whole;
    const shd = '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>';
    const tcPrMatch = inner.match(/(<w:tcPr[\s\S]*?)<\/w:tcPr>/);
    if (tcPrMatch) {
      // 已有 tcPr：在其尾端後插入 shading
      const injected = inner.replace(/(<w:tcPr[\s\S]*?)(<\/w:tcPr>)/, '$1' + shd + '$2');
      return open + injected + close;
    }
    // 沒有 tcPr：在 tc 開標籤後直接插入 tcPr
    return open + '<w:tcPr>' + shd + '</w:tcPr>' + inner + close;
  });
}

function buildClassDict(classCode, yearNum, semNum) {
  const info = idx.classByCode[classCode] || {};
  const className = info['班級名稱'] || classCode;
  const homeTeacher = resolveHomeTeacher(classCode, info);
  const courses = collectClassCourses(classCode);
  const dict = {
    '年': yearNum,
    '期': semNum,
    '班級名稱': className,
    '導師': homeTeacher
  };

  for (let d = 1; d <= 5; d++) {
    dict['d' + d + 'p0'] = slotSubject(classCode, d, WORD_EARLY_PERIOD);
    for (let p = 1; p <= 7; p++) {
      dict['d' + d + 'p' + p] = slotSubject(classCode, d, p);
    }
    dict['d' + d + 'p45'] = slotSubject(classCode, d, WORD_LUNCH_PERIOD);
    // 第八節：單／雙雙欄 + 合併向後相容
    dict['d' + d + 'p8']    = slotSubject(classCode, d, 8);
    dict['d' + d + 'p8s']   = slotSubjectP8(classCode, d, '單週');
    dict['d' + d + 'p8d']   = slotSubjectP8(classCode, d, '雙週');
  }

  // 課表配色：依每格科目＋班級解析底色（班級課表）
  const fills = {};
  for (let d = 1; d <= 5; d++) {
    const earlyFill = resolveScheduleColor(slotSubject(classCode, d, WORD_EARLY_PERIOD), classCode) || '';
    if (earlyFill) fills['d' + d + 'p0'] = earlyFill;
    for (let p = 1; p <= 8; p++) {
      const fill = resolveScheduleColor(slotSubject(classCode, d, p), classCode) || '';
      if (fill) fills['d' + d + 'p' + p] = fill;
    }
    const lunchFill = resolveScheduleColor(slotSubject(classCode, d, WORD_LUNCH_PERIOD), classCode) || '';
    if (lunchFill) fills['d' + d + 'p45'] = lunchFill;
    const fs = resolveScheduleColor(slotSubjectP8(classCode, d, '單週'), classCode) || '';
    const fd = resolveScheduleColor(slotSubjectP8(classCode, d, '雙週'), classCode) || '';
    if (fs) fills['d' + d + 'p8s'] = fs;
    if (fd) fills['d' + d + 'p8d'] = fd;
  }
  dict.__fills = fills;

  FIXED_SLOT_NAMES.forEach(name => {
    let data = findCourseFuzzy(courses, name);
    if (!data && name === '生物') {
      data = findCourseFuzzy(courses, '理化') || findCourseFuzzy(courses, '自然') || findCourseFuzzy(courses, '地球科學');
    }
    dict[name + '節'] = data ? String(data.periods || '') : '';
    dict[name + '師'] = data ? (data.teachers || []).join('\n') : '';
  });
  const bio = findCourseFuzzy(courses, '生物') || findCourseFuzzy(courses, '理化') || findCourseFuzzy(courses, '自然') || findCourseFuzzy(courses, '地球科學');
  dict['生物名'] = bio ? bio.subject : '生物';

  const flexItems = collectFlexItems(courses);
  for (let i = 1; i <= 6; i++) {
    const data = flexItems[i - 1];
    dict['f' + i + '科'] = data ? data.subject : '';
    dict['f' + i + '節'] = data ? String(data.periods || '') : '';
    dict['f' + i + '師'] = data ? (data.teachers || []).join('\n') : '';
  }
  if (flexItems.length > 6) {
    console.warn('彈性超過 6 格未寫入：', flexItems.slice(6).map(x => x.subject).join('、'));
  }
  dict.__flexCount = Math.min(flexItems.length, 6);

  return dict;
}

function collectFlexItems(courses) {
  const flexItems = [];
  const seen = new Set();
  function addFlex(data) {
    if (!data || !data.subject) return;
    if (!isFlexSubject(data.subject)) return;
    const id = normalizeSubjectKey(data.subject) + '|' + (data.teachers || []).join(',');
    if (seen.has(id)) return;
    seen.add(id);
    flexItems.push(data);
  }
  FLEX_SUBJECT_ORDER.forEach(n => addFlex(findCourseFuzzy(courses, n)));
  Object.keys(courses).forEach(k => addFlex(courses[k]));
  flexItems.sort((a, b) => {
    const rank = (d) => {
      const bare = stripSubjectNoise(d.subject);
      const i = FLEX_SUBJECT_ORDER.findIndex(o => bare === o || bare.startsWith(o) || o.startsWith(bare));
      return i < 0 ? 1000 : i;
    };
    return rank(a) - rank(b) || String(a.subject).localeCompare(String(b.subject), 'zh-Hant');
  });
  return flexItems;
}

function mergeEmptyFlexRows(pageXml, flexCount) {
  const n = Math.max(0, Math.min(6, flexCount | 0));
  if (n >= 6) return pageXml;

  const tblRe = /<w:tbl[ >][\s\S]*?<\/w:tbl>/;
  const m = pageXml.match(tblRe);
  if (!m) return pageXml;
  const tbl = m[0];
  const rows = tbl.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g);
  if (!rows || rows.length < 21) return pageXml;

  const FLEX_START = 15;
  const lastFilled = n === 0 ? 14 : (FLEX_START + n - 1);
  const emptyFrom = n === 0 ? FLEX_START : (FLEX_START + n);
  const emptyTo = 20;

  function splitCells(rowXml) {
    return rowXml.match(/<w:tc[ >][\s\S]*?<\/w:tc>/g) || [];
  }
  function rebuildRow(rowXml, cells) {
    const open = rowXml.match(/^<w:tr[^>]*>/)[0];
    const trPr = (rowXml.match(/<w:trPr[\s\S]*?<\/w:trPr>/) || [''])[0];
    return open + trPr + cells.join('') + '</w:tr>';
  }
  function setVMerge(tcXml, mode) {
    let tcPr = (tcXml.match(/<w:tcPr[\s\S]*?<\/w:tcPr>/) || [null])[0];
    if (!tcPr) {
      tcPr = '<w:tcPr></w:tcPr>';
      tcXml = tcXml.replace(/^<w:tc([^>]*)>/, '<w:tc$1>' + tcPr);
    }
    tcPr = tcPr.replace(/<w:vMerge[^/]*\/>/g, '').replace(/<w:vMerge[\s\S]*?<\/w:vMerge>/g, '');
    if (mode === 'restart') {
      tcPr = tcPr.replace(/<\/w:tcPr>/, '<w:vMerge w:val="restart"/></w:tcPr>');
    } else if (mode === 'continue') {
      tcPr = tcPr.replace(/<\/w:tcPr>/, '<w:vMerge/></w:tcPr>');
    }
    if (/<w:tcPr[\s\S]*?<\/w:tcPr>/.test(tcXml)) {
      return tcXml.replace(/<w:tcPr[\s\S]*?<\/w:tcPr>/, tcPr);
    }
    return tcXml.replace(/^<w:tc([^>]*)>/, '<w:tc$1>' + tcPr);
  }
  function flexCellIdxs(cells) {
    if (cells.length >= 12) return [9, 10, 11];
    if (cells.length >= 11) return [9, 10];
    if (cells.length >= 10) return [9];
    return [];
  }

  const newRows = rows.slice();
  {
    const cells = splitCells(newRows[lastFilled]);
    const idxs = flexCellIdxs(cells);
    idxs.forEach(i => {
      if (cells[i]) cells[i] = setVMerge(cells[i], 'restart');
    });
    newRows[lastFilled] = rebuildRow(newRows[lastFilled], cells);
  }
  for (let ri = emptyFrom; ri <= emptyTo; ri++) {
    const cells = splitCells(newRows[ri]);
    const idxs = flexCellIdxs(cells);
    idxs.forEach(i => {
      if (!cells[i]) return;
      let c = setVMerge(cells[i], 'continue');
      c = c.replace(/<w:t([^>]*)>[^<]*<\/w:t>/g, '<w:t$1></w:t>');
      cells[i] = c;
    });
    newRows[ri] = rebuildRow(newRows[ri], cells);
  }

  const open = tbl.match(/^<w:tbl[^>]*>/)[0];
  const tblPr = (tbl.match(/<w:tblPr[\s\S]*?<\/w:tblPr>/) || [''])[0];
  const tblGrid = (tbl.match(/<w:tblGrid[\s\S]*?<\/w:tblGrid>/) || [''])[0];
  const newTbl = open + tblPr + tblGrid + newRows.join('') + '</w:tbl>';
  return pageXml.replace(tblRe, newTbl);
}

async function loadTemplate() {
  if (_tplCache) return _tplCache;
  const resp = await fetch('class-official-template.docx?t=' + Date.now());
  if (!resp.ok) throw new Error('無法載入模板 HTTP ' + resp.status);
  const buf = await resp.arrayBuffer();
  const zip = parseDocxZip(buf);
  const docXml = zip.file('word/document.xml').asText();
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) throw new Error('模板缺少 w:body');
  let bodyInner = bodyMatch[1];
  const sectPr = (bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/) || [''])[0];
  bodyInner = bodyInner.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/, '');
  bodyInner = bodyInner.replace(/(<w:p[ >][\s\S]*?<\/w:p>)\s*$/, (m) => {
    const txt = m.replace(/<[^>]+>/g, '').trim();
    return (txt || /w:br/.test(m)) ? m : '';
  });
  bodyInner = bodyInner.replace(/<w:lastRenderedPageBreak\s*\/>/g, '');
  bodyInner = joinSplitPlaceholders(bodyInner);
  const nextPageSectPr = sectPr
    ? sectPr.replace('</w:sectPr>', '<w:type w:val="nextPage"/></w:sectPr>')
    : '';

  _tplCache = { buf, docXml, bodyInner, sectPr, nextPageSectPr };
  return _tplCache;
}

async function loadTeacherTemplate() {
  if (_teacherTplCache) return _teacherTplCache;
  const resp = await fetch('teacher-official-template.docx?t=' + Date.now());
  if (!resp.ok) throw new Error('無法載入教師模板 HTTP ' + resp.status);
  const buf = await resp.arrayBuffer();
  const zip = parseDocxZip(buf);
  const docXml = zip.file('word/document.xml').asText();
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) throw new Error('教師模板缺少 w:body');
  let bodyInner = bodyMatch[1];
  const sectPr = (bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/) || [''])[0];
  bodyInner = bodyInner.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/, '');
  bodyInner = bodyInner.replace(/(<w:p[ >][\s\S]*?<\/w:p>)\s*$/, (m) => {
    const txt = m.replace(/<[^>]+>/g, '').trim();
    return (txt || /w:br/.test(m)) ? m : '';
  });
  bodyInner = bodyInner.replace(/<w:lastRenderedPageBreak\s*\/>/g, '');
  bodyInner = joinSplitPlaceholders(bodyInner);
  const nextPageSectPr = sectPr
    ? sectPr.replace('</w:sectPr>', '<w:type w:val="nextPage"/></w:sectPr>')
    : '';

  _teacherTplCache = { buf, docXml, bodyInner, sectPr, nextPageSectPr };
  return _teacherTplCache;
}

function injectPageBreakAtStart(pageXml) {
  let out = pageXml;
  out = out.replace(/(<w:spacing[^>]*?)\s+w:before="\d+"/, '$1 w:before="0"');
  out = out.replace('</w:pPr>', '</w:pPr><w:r><w:br w:type="page"/></w:r>');
  return out;
}

function buildClassPageXml(tpl, classCode, yearNum, semNum, leadPageBreak) {
  const dict = buildClassDict(classCode, yearNum, semNum);
  const flexCount = dict.__flexCount || 0;
  delete dict.__flexCount;
  const fills = dict.__fills || null;
  delete dict.__fills;
  let page = expandWordSpecialRows(tpl.bodyInner, dict, 'class');
  page = fills ? injectCellFills(page, fills) : page;
  page = fillPlaceholders(page, dict);
  if (leadPageBreak) page = injectPageBreakAtStart(page);
  return page;
}

function formatClassRanges(classList) {
  if (!classList || classList.length === 0) return '';
  const sorted = [...classList].sort((a, b) => String(a).localeCompare(String(b), 'zh-Hant', { numeric: true }));
  if (sorted.length <= 1) return sorted.join('');

  const groups = [];
  let currentGroup = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    const prev = sorted[i - 1];
    const curr = sorted[i];
    const prevNum = parseInt(prev.replace(/\D/g, ''), 10);
    const currNum = parseInt(curr.replace(/\D/g, ''), 10);
    const prevPrefix = prev.replace(/\d/g, '');
    const currPrefix = curr.replace(/\d/g, '');

    if (!isNaN(prevNum) && !isNaN(currNum) && currNum === prevNum + 1 && prevPrefix === currPrefix) {
      currentGroup.push(curr);
    } else {
      groups.push(currentGroup);
      currentGroup = [curr];
    }
  }
  groups.push(currentGroup);

  return groups.map(g => {
    if (g.length >= 2) return `${g[0]}-${g[g.length - 1]}`;
    return g[0];
  }).join('、');
}

function collectTeacherCourseSummary(teacherCode) {
  const subMap = {};

  (state.schedule || []).forEach(s => {
    // 多師格「教師代碼」為 JSON 字串，需以 getCellTeacherList 解析後比對教師
    const tList = getCellTeacherList(s);
    if (!tList.some(t => String(t['教師姓名'] || t['姓名'] || '').trim() === String(teacherCode))) return;
    const sub = String(s['科目代碼'] || '').trim();
    const clsCode = String(s['班級代碼'] || '').trim();
    if (!sub || !clsCode) return;

    if (!subMap[sub]) subMap[sub] = { classes: new Set(), periodCount: 0, hasOvertime: false };
    subMap[sub].classes.add(clsCode);
    subMap[sub].periodCount++;
    if (isWordOvertimeCell(s)) subMap[sub].hasOvertime = true;
  });

  const summaryList = [];
  Object.keys(subMap).forEach(sub => {
    const item = subMap[sub];
    const classArr = Array.from(item.classes);
    const classRange = formatClassRanges(classArr);
    const total = item.periodCount;
    const countPerClass = classArr.length > 0 ? (total / classArr.length) : 0;
    
    let periodText = '';
    if (classArr.length > 1 && Number.isInteger(countPerClass)) {
      periodText = `各${countPerClass}`;
    } else {
      periodText = `${total}`;
    }

    summaryList.push({
      subject: sub + (item.hasOvertime ? '（超）' : ''),
      subjectCode: sub,
      classRange: classRange,
      hours: periodText,
      periodCount: total
    });
  });

  // 依夫子指示排序：部定課程在前、彈性課程在後，多節在前
  summaryList.sort((a, b) => {
    const aIsBuding = isBudingSubject(a.subjectCode || a.subject);
    const bIsBuding = isBudingSubject(b.subjectCode || b.subject);

    // 1. 部定課程 (true) 在前、彈性課程 (false) 在後
    if (aIsBuding !== bIsBuding) {
      return aIsBuding ? -1 : 1;
    }

    // 2. 多節在前 (periodCount 越大越靠前)
    if (b.periodCount !== a.periodCount) {
      return b.periodCount - a.periodCount;
    }

    // 3. 科目名稱國語排序
    return String(a.subjectCode || a.subject).localeCompare(String(b.subjectCode || b.subject), 'zh-Hant');
  });

  return summaryList;
}

function buildTeacherDict(teacherCode, yearNum, semNum) {
  const tInfo = idx.teacherByCode[teacherCode] || {};
  const name = tInfo['姓名'] || teacherCode;
  const dict = {
    '年': yearNum,
    '期': semNum === '1' ? '一' : (semNum === '2' ? '二' : semNum),
    '姓名': name
  };

  // 課表 grid (Table 0)；同節跨班（協同/合班上課）時合併顯示所有班級
  for (let d = 1; d <= 5; d++) {
    const earlyCells = idx.schedByTeacherSlot[teacherCode + '|' + d + '|' + WORD_EARLY_PERIOD] || [];
    dict[`d${d}p0_s`] = teacherWordSpecialSubject(earlyCells);
    dict[`d${d}p0_c`] = '';
    for (let p = 1; p <= 8; p++) {
      const tk = teacherCode + '|' + d + '|' + p;
      const cells = idx.schedByTeacherSlot[tk] || [];
      if (cells.length > 0) {
        const subLabels = [...new Set(cells.map(c => teacherWordSubject(c)).filter(Boolean))];
        const clsCodes = [...new Set(cells.map(c => String(c['班級代碼'] || '').trim()).filter(Boolean))];
        dict[`d${d}p${p}_s`] = subLabels.join(' / ');
        // 依夫子指示：班級代號（如 701, 802），多班以「 / 」分隔
        dict[`d${d}p${p}_c`] = '';
      } else {
        dict[`d${d}p${p}_s`] = '';
        dict[`d${d}p${p}_c`] = '';
      }
    }
    const lunchCells = idx.schedByTeacherSlot[teacherCode + '|' + d + '|' + WORD_LUNCH_PERIOD] || [];
    dict[`d${d}p45_s`] = teacherWordSpecialSubject(lunchCells);
    dict[`d${d}p45_c`] = '';
  }

  // 配課總表 (Table 1)
  const summaries = collectTeacherCourseSummary(teacherCode);
  for (let i = 1; i <= 6; i++) {
    const rowIdx = Math.ceil(i / 2);
    const colSide = (i % 2 === 1) ? 1 : 2;
    const item = summaries[i - 1];

    dict[`t${rowIdx}_s${colSide}`] = item ? item.subject : '';
    dict[`t${rowIdx}_c${colSide}`] = item ? item.classRange : '';
    dict[`t${rowIdx}_h${colSide}`] = item ? item.hours : '';
  }

  // 課表配色：教師課表格以該格科目為主（多班取首班判斷班級規則）；同格班列一起上色
  const fills = {};
  for (let d = 1; d <= 5; d++) {
    const earlyCells = idx.schedByTeacherSlot[teacherCode + '|' + d + '|' + WORD_EARLY_PERIOD] || [];
    if (earlyCells.length > 0) {
      const fill = resolveScheduleColor(String(earlyCells[0]['科目代碼'] || '').trim(), String(earlyCells[0]['班級代碼'] || '').trim()) || '';
      if (fill) fills[`d${d}p0_s`] = fill;
    }
    for (let p = 1; p <= 8; p++) {
      const tk = teacherCode + '|' + d + '|' + p;
      const cells = idx.schedByTeacherSlot[tk] || [];
      if (cells.length === 0) continue;
      const sub = String(cells[0]['科目代碼'] || '').trim();
      const cls = String(cells[0]['班級代碼'] || '').trim();
      const fill = resolveScheduleColor(sub, cls) || '';
      if (fill) {
        fills[`d${d}p${p}_s`] = fill;
      }
    }
    const lunchCells = idx.schedByTeacherSlot[teacherCode + '|' + d + '|' + WORD_LUNCH_PERIOD] || [];
    if (lunchCells.length > 0) {
      const fill = resolveScheduleColor(String(lunchCells[0]['科目代碼'] || '').trim(), String(lunchCells[0]['班級代碼'] || '').trim()) || '';
      if (fill) fills[`d${d}p45_s`] = fill;
    }
  }
  dict.__fills = fills;

  return dict;
}

function buildTeacherPageXml(tpl, teacherCode, yearNum, semNum, leadPageBreak) {
  const dict = buildTeacherDict(teacherCode, yearNum, semNum);
  const fills = dict.__fills || null;
  delete dict.__fills;
  let page = expandWordSpecialRows(tpl.bodyInner, dict, 'teacher');
  page = fills ? injectCellFills(page, fills) : page;
  page = fillPlaceholders(page, dict);
  if (leadPageBreak) page = injectPageBreakAtStart(page);
  return page;
}

function buildPatrolRoomDict(yearNum, semNum) {
  const dayNames = ['', '週一', '週二', '週三', '週四', '週五'];
  const patrolRows = (state.schedule || [])
    .filter(isPatrolScheduleEntry)
    .sort((a, b) => parseInt(a['星期'], 10) - parseInt(b['星期'], 10) ||
      parseInt(a['節次'], 10) - parseInt(b['節次'], 10) ||
      String(a['教師姓名'] || '').localeCompare(String(b['教師姓名'] || ''), 'zh-Hant'));
  const dict = {
    '年': yearNum,
    '期': semNum === '1' ? '一' : (semNum === '2' ? '二' : semNum),
    '教室': '全校巡堂'
  };
  for (let day = 1; day <= 5; day++) {
    for (let period = 1; period <= 8; period++) {
      const rows = patrolRows.filter(row => parseInt(row['星期'], 10) === day && parseInt(row['節次'], 10) === period);
      const teachers = rows.map(row => teacherName(String(row['教師姓名'] || '').trim())).filter(Boolean);
      dict[`d${day}p${period}_s`] = teachers.length ? '巡堂' : '';
      dict[`d${day}p${period}_c`] = teachers.join('／');
    }
  }
  patrolRows.slice(0, 12).forEach((row, index) => {
    const rowIndex = Math.ceil((index + 1) / 2);
    const side = ((index + 1) % 2 === 1) ? 1 : 2;
    const code = String(row['教師姓名'] || '').trim();
    dict[`t${rowIndex}_s${side}`] = '巡堂';
    dict[`t${rowIndex}_c${side}`] = (dayNames[parseInt(row['星期'], 10)] || '星期' + row['星期']) + '第' + row['節次'] + '節';
    dict[`t${rowIndex}_h${side}`] = teacherName(code);
  });
  return dict;
}

function buildPatrolRoomPageXml(tpl, yearNum, semNum) {
  const filled = fillPlaceholders(tpl.bodyInner, buildPatrolRoomDict(yearNum, semNum));
  const tableStarts = [...filled.matchAll(/<w:tbl(?:\s[^>]*)?>/g)].map(match => match.index);
  let output = filled;
  if (tableStarts.length) {
    const start = tableStarts[tableStarts.length - 1];
    const end = output.indexOf('</w:tbl>', start);
    if (end >= 0) output = output.slice(0, start) + output.slice(end + '</w:tbl>'.length);
  }
  const heading = '任課班級、科目與教師';
  const headingIndex = output.indexOf(heading);
  if (headingIndex >= 0) {
    const paragraphStarts = [...output.matchAll(/<w:p(?:\s[^>]*)?>/g)]
      .map(match => match.index)
      .filter(index => index < headingIndex);
    const paragraphStart = paragraphStarts[paragraphStarts.length - 1];
    const paragraphEnd = output.indexOf('</w:p>', headingIndex);
    if (paragraphStart !== undefined && paragraphEnd >= 0) {
      output = output.slice(0, paragraphStart) + output.slice(paragraphEnd + '</w:p>'.length);
    }
  }
  return output;
}

async function startPatrolWordExport() {
  const patrolCount = (state.schedule || []).filter(isPatrolScheduleEntry).length;
  if (patrolCount === 0) {
    toast('目前沒有巡堂時段可匯出', 'warning');
    return;
  }
  const btn = document.getElementById('word-export-btn');
  const progress = document.getElementById('word-export-progress');
  const msg = document.getElementById('word-export-msg');
  const bar = document.getElementById('word-export-bar');
  if (btn) btn.disabled = true;
  if (progress) progress.style.display = 'block';
  if (msg) msg.textContent = '載入 Word 模板…';
  if (bar) bar.style.width = '20%';
  try {
    _roomTplCache = null;
    const tpl = await loadRoomTemplate();
    const settingsMap = state.settings || {};
    const termCode = settingsMap['學期代號'] || '114-1';
    const yearNum = termCode.split('-')[0] || '114';
    const semNum = termCode.split('-')[1] || '1';
    if (msg) msg.textContent = '整理巡堂資料…';
    if (bar) bar.style.width = '65%';
    const bodyInner = buildPatrolRoomPageXml(tpl, yearNum, semNum) + (tpl.sectPr || '');
    const zip = parseDocxZip(tpl.buf);
    let docXml = zip.file('word/document.xml').asText();
    docXml = docXml.replace(/<w:body[^>]*>[\s\S]*<\/w:body>/, '<w:body>' + bodyInner + '</w:body>');
    zip.file('word/document.xml', docXml);
    if (msg) msg.textContent = '寫入 Word…';
    if (bar) bar.style.width = '90%';
    const blob = zip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    saveAs(blob, `${yearNum}學年度第${semNum}學期巡堂表.docx`);
    if (bar) bar.style.width = '100%';
    if (msg) msg.textContent = '✅ 完成！共 ' + patrolCount + ' 節';
    if (btn) {
      btn.disabled = false;
      btn.textContent = '✅ 完成';
    }
    toast('已匯出巡堂表，共 ' + patrolCount + ' 節', 'success');
  } catch (e) {
    console.error(e);
    toast('❌ 巡堂表匯出失敗：' + (e.message || e), 'error');
    if (btn) btn.disabled = false;
    if (progress) progress.style.display = 'none';
  }
}

async function startClassWordExport() {
  const selected = [...document.querySelectorAll('.word-cls-chk:checked')].map(c => c.value);
  if (selected.length === 0) {
    toast('請至少勾選一個班級', 'warning');
    return;
  }

  const btn = document.getElementById('word-export-btn');
  const progress = document.getElementById('word-export-progress');
  const msg = document.getElementById('word-export-msg');
  const bar = document.getElementById('word-export-bar');

  if (btn) btn.disabled = true;
  if (progress) progress.style.display = 'block';
  if (msg) msg.textContent = '載入模板…';
  if (bar) bar.style.width = '5%';

  try {
    _tplCache = null;
    const tpl = await loadTemplate();
    if (bar) bar.style.width = '12%';

    const settingsMap = state.settings || {};
    const termCode = settingsMap['學期代號'] || '114-1';
    const yearNum = termCode.split('-')[0] || '114';
    const semNum = termCode.split('-')[1] || '1';

    const pages = [];
    for (let i = 0; i < selected.length; i++) {
      const classCode = selected[i];
      if (!idx.classByCode[classCode]) continue;
      const name = idx.classByCode[classCode]['班級名稱'] || classCode;
      if (msg) msg.textContent = `正在填入 ${name}（${i + 1}/${selected.length}）…`;
      if (bar) bar.style.width = `${12 + ((i + 1) / selected.length) * 78}%`;
      pages.push(buildClassPageXml(tpl, classCode, yearNum, semNum, i > 0));
      await new Promise(r => setTimeout(r, 0));
    }

    if (pages.length === 0) {
      toast('❌ 沒有可匯出的班級', 'error');
      if (btn) btn.disabled = false;
      if (progress) progress.style.display = 'none';
      return;
    }

    if (msg) msg.textContent = '寫入 Word…';
    if (bar) bar.style.width = '95%';

    const bodyInner = pages.join('') + (tpl.sectPr || '');
    const zip = parseDocxZip(tpl.buf);
    let docXml = zip.file('word/document.xml').asText();
    docXml = docXml.replace(/<w:body[^>]*>[\s\S]*<\/w:body>/, '<w:body>' + bodyInner + '</w:body>');
    zip.file('word/document.xml', docXml);

    const blob = zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    saveAs(blob, `${yearNum}學年度第${semNum}學期班級課表.docx`);

    if (bar) bar.style.width = '100%';
    if (msg) msg.textContent = `✅ 完成！共 ${pages.length} 班`;
    if (btn) {
      btn.disabled = false;
      btn.textContent = '✅ 完成';
    }
    toast(`已匯出 ${pages.length} 班（佔位符替換，樣式全用模板）`, 'success');
  } catch (e) {
    console.error(e);
    toast('❌ 匯出失敗：' + (e.message || e), 'error');
    if (btn) btn.disabled = false;
    if (progress) progress.style.display = 'none';
  }
}

async function startTeacherWordExport() {
  const selected = [...document.querySelectorAll('.word-t-chk:checked')].map(c => c.value);
  if (selected.length === 0) {
    toast('請至少勾選一位教師', 'warning');
    return;
  }

  const btn = document.getElementById('word-export-btn');
  const progress = document.getElementById('word-export-progress');
  const msg = document.getElementById('word-export-msg');
  const bar = document.getElementById('word-export-bar');

  if (btn) btn.disabled = true;
  if (progress) progress.style.display = 'block';
  if (msg) msg.textContent = '載入教師模板…';
  if (bar) bar.style.width = '5%';

  try {
    _teacherTplCache = null;
    const tpl = await loadTeacherTemplate();
    if (bar) bar.style.width = '12%';

    const settingsMap = state.settings || {};
    const termCode = settingsMap['學期代號'] || '114-1';
    const yearNum = termCode.split('-')[0] || '114';
    const semNum = termCode.split('-')[1] || '1';

    const pages = [];
    for (let i = 0; i < selected.length; i++) {
      const tcCode = selected[i];
      const tInfo = idx.teacherByCode[tcCode] || {};
      const name = tInfo['姓名'] || tcCode;

      if (msg) msg.textContent = `正在填入 ${name} 教師（${i + 1}/${selected.length}）…`;
      if (bar) bar.style.width = `${12 + ((i + 1) / selected.length) * 78}%`;
      pages.push(buildTeacherPageXml(tpl, tcCode, yearNum, semNum, i > 0));
      await new Promise(r => setTimeout(r, 0));
    }

    if (pages.length === 0) {
      toast('❌ 沒有可匯出的教師', 'error');
      if (btn) btn.disabled = false;
      if (progress) progress.style.display = 'none';
      return;
    }

    if (msg) msg.textContent = '寫入 Word…';
    if (bar) bar.style.width = '95%';

    const bodyInner = pages.join('') + (tpl.sectPr || '');
    const zip = parseDocxZip(tpl.buf);
    let docXml = zip.file('word/document.xml').asText();
    docXml = docXml.replace(/<w:body[^>]*>[\s\S]*<\/w:body>/, '<w:body>' + bodyInner + '</w:body>');
    zip.file('word/document.xml', docXml);

    const blob = zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    saveAs(blob, `${yearNum}學年度第${semNum}學期教師課表.docx`);

    if (bar) bar.style.width = '100%';
    if (msg) msg.textContent = `✅ 完成！共 ${pages.length} 位教師`;
    if (btn) {
      btn.disabled = false;
      btn.textContent = '✅ 完成';
    }
    toast(`已匯出 ${pages.length} 位教師課表（含配課總表）`, 'success');
  } catch (e) {
    console.error(e);
    toast('❌ 匯出失敗：' + (e.message || e), 'error');
    if (btn) btn.disabled = false;
    if (progress) progress.style.display = 'none';
  }
}

async function loadRoomTemplate() {
  if (_roomTplCache) return _roomTplCache;
  const resp = await fetch('room-official-template.docx?t=' + Date.now());
  if (!resp.ok) throw new Error('無法載入教室模板 HTTP ' + resp.status);
  const buf = await resp.arrayBuffer();
  const zip = parseDocxZip(buf);
  const docXml = zip.file('word/document.xml').asText();
  const bodyMatch = docXml.match(/<w:body[^>]*>([\s\S]*)<\/w:body>/);
  if (!bodyMatch) throw new Error('教室模板缺少 w:body');
  let bodyInner = bodyMatch[1];
  const sectPr = (bodyInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/) || [''])[0];
  bodyInner = bodyInner.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/, '');
  bodyInner = bodyInner.replace(/(<w:p[ >][\s\S]*?<\/w:p>)\s*$/, (m) => {
    const txt = m.replace(/<[^>]+>/g, '').trim();
    return (txt || /w:br/.test(m)) ? m : '';
  });
  bodyInner = bodyInner.replace(/<w:lastRenderedPageBreak\s*\/>/g, '');
  bodyInner = joinSplitPlaceholders(bodyInner);
  const nextPageSectPr = sectPr
    ? sectPr.replace('</w:sectPr>', '<w:type w:val="nextPage"/></w:sectPr>')
    : '';

  _roomTplCache = { buf, docXml, bodyInner, sectPr, nextPageSectPr };
  return _roomTplCache;
}

function collectRoomTeacherSummary(roomCode) {
  const teacherMap = {};

  (state.schedule || []).forEach(s => {
    const subCode = String(s['科目代碼'] || '').trim();
    if (!subCode) return;
    const subInfo = idx.subjectByCode[subCode];
    const room = subInfo ? String(subInfo['所屬教室代碼'] || '').trim() : '';
    if (room !== String(roomCode)) return;

    const clsCode = String(s['班級代碼'] || '').trim();
    const tcCode  = String(s['教師姓名'] || '').trim();
    if (!tcCode) return;

    if (!teacherMap[tcCode]) {
      teacherMap[tcCode] = {
        teacherCode: tcCode,
        teacherName: teacherName(tcCode),
        subjects: new Set(),
        classes: new Set(),
        periodCount: 0
      };
    }
    if (subCode) teacherMap[tcCode].subjects.add(subCode);
    if (clsCode) teacherMap[tcCode].classes.add(clsCode);
    teacherMap[tcCode].periodCount++;
  });

  const summaryList = [];
  Object.keys(teacherMap).forEach(tcCode => {
    const item = teacherMap[tcCode];
    const classArr = Array.from(item.classes).sort();
    const classRange = formatClassRanges(classArr);
    const subStr = Array.from(item.subjects).join('、');

    summaryList.push({
      teacher: item.teacherName || tcCode,
      subject: subStr,
      classRange: classRange,
      periodCount: item.periodCount
    });
  });

  summaryList.sort((a, b) => {
    if (b.periodCount !== a.periodCount) return b.periodCount - a.periodCount;
    return String(a.teacher).localeCompare(String(b.teacher), 'zh-Hant');
  });

  return summaryList;
}

function buildRoomDict(roomCode, yearNum, semNum) {
  const roomObj = idx.roomByCode[roomCode] || {};
  const roomName = roomObj['教室名稱'] || roomCode;

  const dict = {
    '年': yearNum,
    '期': semNum === '1' ? '一' : (semNum === '2' ? '二' : semNum),
    '教室': roomName
  };

  // 課表 grid (只寫班級代碼，同科目去重不重複顯示)
  for (let d = 1; d <= 5; d++) {
    const earlyCells = idx.schedByRoomSlot[roomCode + '|' + d + '|' + WORD_EARLY_PERIOD] || [];
    const earlySubjects = [...new Set(earlyCells.map(c => String(c['科目代碼'] || '').trim()).filter(Boolean))];
    dict[`d${d}p0_s`] = earlySubjects.join(' / ');
    dict[`d${d}p0_c`] = earlyCells.map(c => String(c['班級代碼'] || '').trim()).filter(Boolean).join(' / ');
    for (let p = 1; p <= 8; p++) {
      const rk = roomCode + '|' + d + '|' + p;
      const cells = idx.schedByRoomSlot[rk] || [];
      if (cells.length > 0) {
        const subCodes = [...new Set(cells.map(c => String(c['科目代碼'] || '').trim()).filter(Boolean))];
        dict[`d${d}p${p}_s`] = subCodes.join(' / ');
        dict[`d${d}p${p}_c`] = cells.map(c => String(c['班級代碼'] || '').trim()).join(' / ');
      } else {
        dict[`d${d}p${p}_s`] = '';
        dict[`d${d}p${p}_c`] = '';
      }
    }
    const lunchCells = idx.schedByRoomSlot[roomCode + '|' + d + '|' + WORD_LUNCH_PERIOD] || [];
    const lunchSubjects = [...new Set(lunchCells.map(c => String(c['科目代碼'] || '').trim()).filter(Boolean))];
    dict[`d${d}p45_s`] = lunchSubjects.join(' / ');
    dict[`d${d}p45_c`] = lunchCells.map(c => String(c['班級代碼'] || '').trim()).filter(Boolean).join(' / ');
  }

  // 配課總表（按教師歸類，同一老師一列/一個位子）
  const summaries = collectRoomTeacherSummary(roomCode);
  for (let i = 1; i <= 6; i++) {
    const rowIdx = Math.ceil(i / 2);
    const colSide = (i % 2 === 1) ? 1 : 2;
    const item = summaries[i - 1];

    dict[`t${rowIdx}_s${colSide}`] = item ? item.subject : '';
    dict[`t${rowIdx}_c${colSide}`] = item ? item.classRange : '';
    dict[`t${rowIdx}_h${colSide}`] = item ? item.teacher : '';
  }

  return dict;
}

function buildRoomPageXml(tpl, roomCode, yearNum, semNum, leadPageBreak) {
  const dict = buildRoomDict(roomCode, yearNum, semNum);
  let page = expandWordSpecialRows(tpl.bodyInner, dict, 'room');
  page = fillPlaceholders(page, dict);
  if (leadPageBreak) page = injectPageBreakAtStart(page);
  return page;
}

async function startRoomWordExport() {
  const selected = [...document.querySelectorAll('.word-r-chk:checked')].map(c => c.value);
  if (selected.length === 0) {
    toast('請至少勾選一間教室', 'warning');
    return;
  }

  const btn = document.getElementById('word-export-btn');
  const progress = document.getElementById('word-export-progress');
  const msg = document.getElementById('word-export-msg');
  const bar = document.getElementById('word-export-bar');

  if (btn) btn.disabled = true;
  if (progress) progress.style.display = 'block';
  if (msg) msg.textContent = '載入教室模板…';
  if (bar) bar.style.width = '5%';

  try {
    _roomTplCache = null;
    const tpl = await loadRoomTemplate();
    if (bar) bar.style.width = '12%';

    const settingsMap = state.settings || {};
    const termCode = settingsMap['學期代號'] || '114-1';
    const yearNum = termCode.split('-')[0] || '114';
    const semNum = termCode.split('-')[1] || '1';

    const pages = [];
    for (let i = 0; i < selected.length; i++) {
      const roomCode = selected[i];
      const roomObj = idx.roomByCode[roomCode] || {};
      const name = roomObj['教室名稱'] || roomCode;

      if (msg) msg.textContent = `正在填入 ${name} 教室（${i + 1}/${selected.length}）…`;
      if (bar) bar.style.width = `${12 + ((i + 1) / selected.length) * 78}%`;
      pages.push(buildRoomPageXml(tpl, roomCode, yearNum, semNum, i > 0));
      await new Promise(r => setTimeout(r, 0));
    }

    if (pages.length === 0) {
      toast('❌ 沒有可匯出的教室', 'error');
      if (btn) btn.disabled = false;
      if (progress) progress.style.display = 'none';
      return;
    }

    if (msg) msg.textContent = '寫入 Word…';
    if (bar) bar.style.width = '95%';

    const bodyInner = pages.join('') + (tpl.sectPr || '');
    const zip = parseDocxZip(tpl.buf);
    let docXml = zip.file('word/document.xml').asText();
    docXml = docXml.replace(/<w:body[^>]*>[\s\S]*<\/w:body>/, '<w:body>' + bodyInner + '</w:body>');
    zip.file('word/document.xml', docXml);

    const blob = zip.generate({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    saveAs(blob, `${yearNum}學年度第${semNum}學期教室課表.docx`);

    if (bar) bar.style.width = '100%';
    if (msg) msg.textContent = `✅ 完成！共 ${pages.length} 間教室`;
    if (btn) {
      btn.disabled = false;
      btn.textContent = '✅ 完成';
    }
    toast(`已匯出 ${pages.length} 間教室課表（含使用總表）`, 'success');
  } catch (e) {
    console.error(e);
    toast('❌ 匯出失敗：' + (e.message || e), 'error');
    if (btn) btn.disabled = false;
    if (progress) progress.style.display = 'none';
  }
}

async function startWordExport() {
  if (_wordCurrentTab === 'class') {
    await startClassWordExport();
  } else if (_wordCurrentTab === 'teacher') {
    await startTeacherWordExport();
  } else if (_wordCurrentTab === 'room') {
    await startRoomWordExport();
  } else {
    await startPatrolWordExport();
  }
}
