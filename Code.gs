/**
 * Журнал курса UI/UX (ПМ 1, РО 1.2) в Google Таблице.
 * Как установить: см. файл «ИНСТРУКЦИЯ.md».
 */

const SHEET_STUDENTS = 'Студенты';
const SHEET_SCHEDULE = 'Расписание';

const MODULES = [
  ['m1', 'Введение в дизайн пользовательских интерфейсов'],
  ['m2', 'Стратегия дизайна интерфейсов'],
  ['m3', 'Основы проектирования взаимодействия (UX)'],
  ['m4', 'Дизайн для разных платформ: веб, мобильные, VR/AR'],
  ['m5', 'Современные тренды и паттерны'],
  ['m6', 'Презентация и защита дизайн-решений']
];
const TASKS = [
  ['l1', 'ЛПР 1'], ['l2', 'ЛПР 2'], ['s1', 'СРСП 1'],
  ['l3', 'ЛПР 3'], ['l4', 'ЛПР 4'], ['s2', 'СРСП 2'],
  ['l5', 'ЛПР 5'], ['l6', 'ЛПР 6'], ['s3', 'СРСП 3'],
  ['l7', 'ЛПР 7'], ['l8', 'ЛПР 8'], ['s4', 'СРСП 4'],
  ['l9', 'ЛПР 9'], ['l10', 'ЛПР 10'], ['s5', 'СРСП 5'],
  ['l11', 'ЛПР 11'], ['l12', 'ЛПР 12'], ['s6', 'СРСП 6']
];
const FIXED = ['Код', 'ФИО', 'Группа', 'Тест 1', 'Тест 2', 'Тест 3', 'Тест 4', 'Тест 5', 'Тест 6',
  'Итоговый тест', 'Дата сертификата', 'Практика отмечена студентом'];
const COL_CODE = 1, COL_NAME = 2, COL_GROUP = 3, COL_TESTS = 4, COL_FINAL = 10, COL_CERT = 11, COL_MARKED = 12;
const COL_GRADES = 13;
const COL_PROGRESS = COL_GRADES + TASKS.length;   // служебный столбец
const COL_UPDATED = COL_PROGRESS + 1;

/* ---------- Меню и настройка ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Курс UI/UX')
    .addItem('1. Подготовить таблицы (один раз)', 'setup')
    .addItem('2. Выдать коды студентам', 'issueCodes')
    .addItem('3. Выдать коды уроков', 'issueLessonCodes')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let st = ss.getSheetByName(SHEET_STUDENTS) || ss.insertSheet(SHEET_STUDENTS);
  const head = FIXED.concat(TASKS.map(t => t[1])).concat(['Прогресс (служебное)', 'Обновлено']);
  st.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#ECEEF7').setWrap(true);
  st.setFrozenRows(1);
  st.setFrozenColumns(3);
  st.hideColumns(COL_PROGRESS);
  st.getRange(1, COL_GRADES, 1, TASKS.length).setBackground('#FFF4D6');

  let sc = ss.getSheetByName(SHEET_SCHEDULE) || ss.insertSheet(SHEET_SCHEDULE);
  if (sc.getLastRow() < 1) {
    sc.getRange(1, 1, 1, 5).setValues([['Модуль', 'Тема', 'Открыть с даты', 'Нужен код (Да/Нет)', 'Код урока']])
      .setFontWeight('bold').setBackground('#ECEEF7');
    const rows = MODULES.map((m, i) => ['Модуль ' + (i + 1), m[1], '', 'Нет', '']);
    sc.getRange(2, 1, rows.length, 5).setValues(rows);
    sc.getRange(2, 3, rows.length, 1).setNumberFormat('yyyy-mm-dd');
    sc.setColumnWidth(2, 380);
    sc.setFrozenRows(1);
  }
  const first = ss.getSheetByName('Лист1') || ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 2 && first.getLastRow() === 0) ss.deleteSheet(first);
  SpreadsheetApp.getUi().alert('Готово. Впишите ФИО и группу на листе «Студенты», затем выберите в меню «Курс UI/UX» пункт «Выдать коды студентам».');
}

function genCode_() {
  const al = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < 6; i++) s += al.charAt(Math.floor(Math.random() * al.length));
  return s;
}

function issueCodes() {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STUDENTS);
  if (!st || st.getLastRow() < 2) { SpreadsheetApp.getUi().alert('Сначала впишите ФИО и группу студентов на листе «Студенты».'); return; }
  const n = st.getLastRow() - 1;
  const vals = st.getRange(2, 1, n, 2).getValues();
  const used = {};
  vals.forEach(r => { const c = String(r[0]).trim(); if (c) used[c] = true; });
  let count = 0;
  vals.forEach((r, i) => {
    if (String(r[1]).trim() && !String(r[0]).trim()) {
      let code;
      do { code = genCode_(); } while (used[code]);
      used[code] = true;
      st.getRange(i + 2, COL_CODE).setValue(code);
      count++;
    }
  });
  SpreadsheetApp.getUi().alert('Выдано кодов: ' + count);
}

function issueLessonCodes() {
  const sc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHEDULE);
  if (!sc) { SpreadsheetApp.getUi().alert('Сначала выполните пункт «Подготовить таблицы».'); return; }
  const n = Math.max(sc.getLastRow() - 1, 0);
  if (!n) return;
  const vals = sc.getRange(2, 1, n, 5).getValues();
  let count = 0;
  vals.forEach((r, i) => {
    if (String(r[3]).trim().toLowerCase() === 'да' && !String(r[4]).trim()) {
      sc.getRange(i + 2, 5).setValue(genCode_());
      count++;
    }
  });
  SpreadsheetApp.getUi().alert('Выдано кодов уроков: ' + count + '. Коды только для строк, где в столбце «Нужен код» стоит «Да».');
}

/* ---------- Веб-приложение ---------- */

function doGet() {
  return json_({ ok: true, message: 'Журнал курса UI/UX работает' });
}

function doPost(e) {
  let out;
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = JSON.parse(e.postData.contents);
    out = route_(p);
  } catch (err) {
    out = { ok: false, error: 'Ошибка сервера. Попробуйте ещё раз.' };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  const code = String(p.code || '').trim().toUpperCase();
  if (!code) return { ok: false, error: 'Введите код.' };
  const found = findStudent_(code);
  if (!found) return { ok: false, error: 'Такого кода нет. Проверьте и введите снова.' };
  const st = found.sheet, row = found.row;

  if (p.action === 'login') {
    return {
      ok: true,
      name: String(st.getRange(row, COL_NAME).getValue()).trim(),
      group: String(st.getRange(row, COL_GROUP).getValue()).trim(),
      progress: parse_(st.getRange(row, COL_PROGRESS).getValue()),
      grades: readGrades_(st, row),
      schedule: readSchedule_()
    };
  }
  if (p.action === 'save') {
    saveProgress_(st, row, p.progress || {});
    return { ok: true };
  }
  if (p.action === 'unlock') {
    return unlock_(st, row, String(p.mid || ''), String(p.lessonCode || ''));
  }
  return { ok: false, error: 'Неизвестное действие.' };
}

function findStudent_(code) {
  const st = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STUDENTS);
  if (!st || st.getLastRow() < 2) return null;
  const codes = st.getRange(2, COL_CODE, st.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0]).trim().toUpperCase() === code) return { sheet: st, row: i + 2 };
  }
  return null;
}

function parse_(v) {
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

function readGrades_(st, row) {
  const vals = st.getRange(row, COL_GRADES, 1, TASKS.length).getValues()[0];
  const out = {};
  TASKS.forEach((t, i) => { if (vals[i] !== '' && vals[i] !== null) out[t[0]] = vals[i]; });
  return out;
}

function readSchedule_() {
  const sc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHEDULE);
  const out = {};
  if (!sc || sc.getLastRow() < 2) return out;
  const tz = Session.getScriptTimeZone();
  const v = sc.getRange(2, 1, Math.min(sc.getLastRow() - 1, MODULES.length), 5).getValues();
  v.forEach((r, i) => {
    let date = '';
    if (r[2] instanceof Date) date = Utilities.formatDate(r[2], tz, 'yyyy-MM-dd');
    else if (r[2]) date = String(r[2]).trim();
    out[MODULES[i][0]] = { date: date, req: String(r[3]).trim().toLowerCase() === 'да' };
  });
  return out;
}

function saveProgress_(st, row, prog) {
  const mods = prog.mods || {};
  const tests = MODULES.map(m => {
    const r = mods[m[0]];
    return r && r.attempts ? r.best : '';
  });
  const fin = prog.final && prog.final.attempts ? prog.final.best : '';
  const cert = prog.final && prog.final.passedAt ? prog.final.passedAt : '';
  const marked = Object.keys(prog.tasks || {}).filter(k => prog.tasks[k]).length;
  const clean = { mods: mods, final: prog.final || null, tasks: prog.tasks || {}, ts: prog.ts || Date.now() };
  st.getRange(row, COL_TESTS, 1, 6).setValues([tests]);
  st.getRange(row, COL_FINAL).setValue(fin);
  st.getRange(row, COL_CERT).setValue(cert);
  st.getRange(row, COL_MARKED).setValue(marked);
  st.getRange(row, COL_PROGRESS).setValue(JSON.stringify(clean));
  st.getRange(row, COL_UPDATED).setValue(new Date());
}

function unlock_(st, row, mid, lessonCode) {
  const sched = readSchedule_();
  const s = sched[mid];
  if (!s) return { ok: false, error: 'Такого модуля нет.' };
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (s.date && today < s.date) return { ok: false, error: 'Эта тема откроется ' + s.date + '.' };
  if (s.req) {
    const sc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHEDULE);
    const idx = MODULES.findIndex(m => m[0] === mid);
    const real = String(sc.getRange(idx + 2, 5).getValue()).trim().toLowerCase();
    if (!real || real !== lessonCode.trim().toLowerCase()) return { ok: false, error: 'Код не подходит. Проверьте и введите снова.' };
  }
  const prog = parse_(st.getRange(row, COL_PROGRESS).getValue()) || { mods: {}, final: null, tasks: {} };
  prog.mods = prog.mods || {};
  prog.mods[mid] = prog.mods[mid] || { best: 0, attempts: 0 };
  prog.mods[mid].unlocked = 'srv';
  st.getRange(row, COL_PROGRESS).setValue(JSON.stringify(prog));
  return { ok: true };
}
