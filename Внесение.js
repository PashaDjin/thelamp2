function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Проводки')
    .addItem('🚀 Провести', 'runTransfer')
    .addSeparator()

    .addItem('Провести выручку по актам', 'createRevenueFromActs')
    .addSeparator()
    .addItem('Провести ЗП', 'createMasterFromActs')
    .addItem('Провести возврат депозитов', 'createDepositReturnFromActs')
    .addToUi();
}
// Константы перемещены в `00-constants.js`
// Утилиты перемещены в `10-utils.js`


/**
 * Показывает HTML-диалог и блокирующе ждёт ответа (до таймаута).
 * @param {Object} options
 * @param {string} options.title
 * @param {string} options.message
 * @param {string[]} options.buttons
 * @param {boolean} [options.withInput]
 * @param {string} [options.defaultValue]
 * @returns {{button: string, value: string}|null}
 */
function showDialogAndWait_({ title, message, buttons, withInput = false, defaultValue = '' }) {
  const cache = CacheService.getScriptCache();
  const token = `dlg_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  cache.remove(token);

  const html = HtmlService.createHtmlOutput(`
    <div style="white-space:pre-wrap;">${escapeHtml_(message)}</div>
    ${withInput ? `<div><input id="dlg-input" value="${escapeHtml_(defaultValue)}" /></div>` : ''}
    <div>${buttons.map(b => `<button onclick="submitDialog('${b}')">${escapeHtml_(b)}</button>`).join('')}</div>
    <script>
      function submitDialog(btn){
        const v = document.getElementById('dlg-input') ? document.getElementById('dlg-input').value : '';
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
          .setDialogResult('${token}', { button: btn, value: v });
      }
      document.addEventListener('DOMContentLoaded', function(){
        const b = document.querySelector('button'); if(b) b.focus();
      });
    </script>
  `)
    .setWidth(380)
    .setHeight(withInput ? 180 : 140);

  SpreadsheetApp.getUi().showModalDialog(html, title);

  const timeoutMs = 20000;
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const data = cache.get(token);
    if (data) {
      cache.remove(token);
      try {
        return JSON.parse(data);
      } catch (e) {
        return null;
      }
    }
    Utilities.sleep(30);
  }

  cache.remove(token);
  return null;
}

// escapeHtml_ moved to `10-utils.js`
function setDialogResult(token, data) {
  // use cache for faster cross-process signalling
  try {
    CacheService.getScriptCache().put(token, JSON.stringify(data || {}), 120);
  } catch (e) {
    // fallback to properties if cache fails for any reason
    PropertiesService.getDocumentProperties().setProperty(token, JSON.stringify(data || {}));
  }
}

function confirmDialog_(title, message) {
  const res = showDialogAndWait_({ title, message, buttons: ['Да', 'Нет'] });
  return !!(res && res.button === 'Да');
}

function okDialog_(title, message) {
  showDialogAndWait_({ title, message, buttons: ['Ок'] });
}

function promptDialog_(title, message, defaultValue) {
  const res = showDialogAndWait_({ title, message, buttons: ['Ок', 'Отмена'], withInput: true, defaultValue });
  if (!res || res.button !== 'Ок') return { button: 'Cancel', text: '' };
  return { button: 'Ok', text: res.value };
}

// Унифицированный ключ акта по адресу и номеру
function makeActKey(addr, actNo) {
  const a = String(addr || '').trim();
  const n = String(actNo || '').trim();
  if (!a && !n) return '';
  return a + '|' + n;
}
function createMasterFromActs() {
  createEntriesFromSelectedActs_({ mode: 'MASTER' });
}

function createDepositReturnFromActs() {
  createEntriesFromSelectedActs_({ mode: 'DEPOSIT_RETURN' });
}

function createRevenueFromActs() {
  createEntriesFromSelectedActs_({ mode: 'REVENUE' });
}

/**
 * Общая логика для:
 *  - mode='MASTER'           → "% Мастер", сумма из J (на руки)
 *  - mode='DEPOSIT_RETURN'   → "Возврат удержания", сумма из I
 *
 * Работает только, если активен лист "РЕЕСТР АКТОВ" и есть выделение.
 */
function createEntriesFromSelectedActs_({ mode }) {
  const ss   = SpreadsheetApp.getActive();
  const shActs = ss.getSheetByName('РЕЕСТР АКТОВ');
  const shIn   = ss.getSheetByName('⏬ ВНЕСЕНИЕ');

  if (!shActs || !shIn) {
    okDialog_('Нет листов', 'Камрад, не нахожу листы "РЕЕСТР АКТОВ" и/или "⏬ ВНЕСЕНИЕ".');
    return;
  }

  // Требуем, чтобы пользователь был на листе "РЕЕСТР АКТОВ"
  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getName() !== shActs.getName()) {
    okDialog_('Не тот лист', 'Камрад, сначала перейди на лист "РЕЕСТР АКТОВ" и выдели строки с актами.');
    return;
  }

  const selection = ss.getSelection();
  const rangeList = selection && selection.getActiveRangeList();
  if (!rangeList) {
    okDialog_('Нет выделения', 'Камрад, выдели хотя бы одну ячейку с актом в "РЕЕСТР АКТОВ".');
    return;
  }

  // Собираем номера строк из всех выделенных диапазонов на "РЕЕСТР АКТОВ"
  const rowSet = new Set();
  rangeList.getRanges().forEach(r => {
    if (r.getSheet().getName() !== shActs.getName()) return;
    const start = r.getRow();
    const end   = r.getLastRow();
    for (let row = start; row <= end; row++) {
      if (row > 1) rowSet.add(row); // выше заголовка не берём
    }
  });

  const rows = Array.from(rowSet).sort((a, b) => a - b);
  if (!rows.length) {
    okDialog_('Пусто', 'Камрад, по выделению не нашёл ни одной строки с актами.');
    return;
  }

  // Читаем данные по каждому акту
  const items = [];
  const errors = [];

  rows.forEach(row => {
    const addr  = shActs.getRange(row, 2).getValue(); // B: адрес
    const actNo = shActs.getRange(row, 3).getValue(); // C: номер акта
    const amountCol =
      mode === 'MASTER'         ? 11 : // J — "на руки"
      mode === 'DEPOSIT_RETURN' ? 10  : // I — возврат депозита
      mode === 'REVENUE'        ? 5  : // E — выручка по акту
      0;

    const amountCell = amountCol ? shActs.getRange(row, amountCol).getValue() : '';
    const amount = Number(amountCell);

    if (!addr || !actNo || amountCell === '' || amountCell == null || !isFinite(amount) || amount === 0) {
      errors.push(`Строка ${row}: пропускаю (нет адреса, акта или суммы).`);
      return;
    }

    items.push({
      row,
      addr: String(addr),
      actNo: String(actNo),
      amount
    });
  });

  if (!items.length) {
    okDialog_('Пусто', 'Камрад, по выбранным строкам нечего проводить (пустые адреса/акты/суммы).');
    return;
  }

  // Подтверждение перед записью
  const title =
    mode === 'MASTER'
      ? 'Провести ЗП мастерам'
      : mode === 'DEPOSIT_RETURN'
        ? 'Провести возврат депозитов'
        : mode === 'REVENUE'
          ? 'Провести выручку по актам'
          : 'Провести операции по актам';

  // Автоматически оформляем проводки по выбранным актам (без HTML-подтверждения).
  const lines = items.map(it => `• ${it.addr} — ${it.amount} (${it.actNo})`);

  // Нормализуем вход (очистка NBSP + trim) в B..F, чтобы ненужные символы не мешали поиску пустой строки
  normalizeInputBF_(shIn);

  // Ищем первую пустую строку во "⏬ ВНЕСЕНИЕ" в блоке B10:F40 (учитываем только B..F)
  const firstRow = findFirstEmptyRowInInput_(shIn);
  if (!firstRow) {
    // Диагностика: выясним, какие именно строки/ячейки заняты в B10:F40 — покажем короткий отчёт
    const diagRange = shIn.getRange(IN_START_ROW, IN_COL_B, IN_HEIGHT, IN_COL_F - IN_COL_B + 1); // B10:F40
    const diagVals  = diagRange.getValues();
    const nonEmptyRows = [];
    for (let ri = 0; ri < diagVals.length; ri++) {
      const row = diagVals[ri];
      const cols = [];
      for (let ci = 0; ci < row.length; ci++) {
        const v = row[ci];
        if (v != null && String(v).trim() !== '') {
          // колонка (B..F)
          const colNum = IN_COL_B + ci;
          const colLetter = String.fromCharCode(64 + colNum);
          let s = String(v);
          s = s.replace(/\n/g, ' ');
          if (s.length > 30) s = s.slice(0, 27) + '...';
          cols.push(`${colLetter}:${s}`);
        }
      }
      if (cols.length) nonEmptyRows.push({row: IN_START_ROW + ri, cols});
    }

    let msg = `Во "⏬ ВНЕСЕНИЕ" нет полностью пустых строк в диапазоне B10:F40 (учитываются только B..F).`;
    msg += '\nНайдено занятых строк: ' + nonEmptyRows.length + '.';
    if (nonEmptyRows.length) {
      msg += '\nПервые несколько (строка: столбцы=значения):\n';
      msg += nonEmptyRows.slice(0, 6).map(r => `• ${r.row}: ${r.cols.join(', ')}`).join('\n');
    }

    okDialog_('Нет места', msg);
    return;
  }

  // === ВАЖНО: учитываем, что при REVENUE на каждый акт будет 2 строки (Выручка + НРП) ===
  const rowsPerItem = (mode === 'REVENUE') ? 2 : 1;
  const lastRowNeeded = firstRow + rowsPerItem * items.length - 1;
  if (lastRowNeeded > 40) {
    okDialog_('Нет места', 'Камрад, не хватает свободных строк во "⏬ ВНЕСЕНИЕ" для всех проводок. Освободи место и попробуй ещё раз.');
    return;
  }

  // Дата по Москве (для MASTER / DEPOSIT_RETURN и для НРП)
  const todayStr  = Utilities.formatDate(new Date(), MOSCOW_TZ, 'dd.MM.yyyy');
  const todayDate = parseSheetDate_(todayStr, MOSCOW_TZ);

  const article =
    mode === 'MASTER'
      ? '% Мастер'
      : mode === 'DEPOSIT_RETURN'
        ? 'Возврат удержания'
        : mode === 'REVENUE'
          ? 'Выручка по акту'
          : '';

  // Готовим массив значений для записи в B..G
  let values = [];

  if (mode === 'REVENUE') {
    // Для выручки по актам: на каждый акт — две строки (Выручка по акту + НРП 3%)
    items.forEach(it => {
      // 1) основная выручка по акту — ставим дату сегодня сразу (как и для НРП)
      // Формат: B дата, C кошелёк, D сумма, E статья, F объект, G акт
      values.push([
        todayDate,       // B: дата — ставим сегодня
        '',              // C: кошелёк
        it.amount,       // D: сумма акта
        article,         // E: "Выручка по акту"
        it.addr,         // F: объект (адрес)
        it.actNo         // G: акт
      ]);

      // 2) НРП — 3% от суммы акта, датой сегодня
      const nrpAmount = Math.round(it.amount * 0.03 * 100) / 100; // округляем до копеек

      values.push([
        todayDate,       // B: дата по Москве
        '',              // C: кошелёк (выберешь сам)
        nrpAmount,       // D: 3% от суммы акта
        'НРП',           // E: статья НРП
        it.addr,         // F: объект
        it.actNo         // G: акт
      ]);
    });
  } else {
    // MASTER / DEPOSIT_RETURN — одна строка на акт; пишем только B..G по просьбе владельца
    values = items.map(it => ([
      todayDate,    // B: дата
      '',           // C: кошелёк
      it.amount,    // D: сумма
      article,      // E: статья
      it.addr,      // F: расшифровка (адрес)
      it.actNo      // G: акт
    ]));
  }

  // Перед записью уверимся, что все строки имеют ровно 6 колонок (B..G).
  // Если кто-то случайно сформировал шире — обрежем, если уже короче — дополним пустыми.
  const EXPECTED_COLS = 6;
  let adjusted = false;
  values = values.map((r, idx) => {
    if (!Array.isArray(r)) {
      adjusted = true;
      return Array(EXPECTED_COLS).fill('');
    }
    if (r.length === EXPECTED_COLS) return r;
    adjusted = true;
    if (r.length > EXPECTED_COLS) return r.slice(0, EXPECTED_COLS);
    return r.concat(Array(EXPECTED_COLS - r.length).fill(''));
  });
  if (adjusted) {
    console.warn('createEntriesFromSelectedActs_: adjusted values rows to width 6 for B..G', values);
    SpreadsheetApp.getActive().toast('Внимание: некоторые строки были приведены к ширине B..G перед записью.', 'Проведение', 6);
  }

  const targetRange = shIn.getRange(firstRow, 2, values.length, EXPECTED_COLS); // B..G
  targetRange.setValues(values);
  // Формат даты для колонки B
  shIn.getRange(firstRow, 2, values.length, 1).setNumberFormat('dd"."mm"."yyyy');

  let msg = `Создано проводок во "⏬ ВНЕСЕНИЕ": ${values.length}.`;
  if (errors.length) {
    msg += `\n\nПропущено строк: ${errors.length}.\nПервые несколько:\n` +
      errors.slice(0, 5).map(e => '• ' + e).join('\n');
  }

  SpreadsheetApp.getActive().toast(msg, 'Готово', 5);

  // Авто-проведение отключено: функция только записывает проводки в "⏬ ВНЕСЕНИЕ".
  // Пользователь может запустить проведение вручную отдельной командой.
}



/**
 * Находит строку начала записи в лист "☑️ ПРОВОДКИ".
 * Использует сохранённый в DocumentProperties номер последней строки,
 * отступает от него 10 строк вверх и ищет первую пустую строку.
 * Если ничего не нашёл — пишет в конец (lastRow + 1).
 */
function findStartRowForProv_(shProv) {
  const props = PropertiesService.getDocumentProperties();
  const hintStr = props.getProperty('LAST_PROV_ROW');
  const lastRow = Math.max(shProv.getLastRow(), 1); // минимум заголовок

  let hint = Number(hintStr);
  if (!Number.isFinite(hint) || hint < 2) {
    // Если подсказки нет или мусор — считаем, что писали в конец
    hint = lastRow;
  }

  // Старт сканирования: на 10 строк выше подсказки, но не выше 2
  let scanFrom = Math.max(2, hint - 10);
  let scanTo   = lastRow;

  if (scanFrom > scanTo) {
    scanFrom = 2;
    scanTo   = lastRow;
  }

  const height = Math.max(1, scanTo - scanFrom + 1);
  const grid = shProv.getRange(scanFrom, 1, height, 10).getValues();

  let start = lastRow + 1; // по умолчанию — в конец

  for (let i = 0; i < grid.length; i++) {
    const row = grid[i];
    const isEmpty = row.every(v => v === '' || v === null);
    if (isEmpty) {
      start = scanFrom + i;
      break;
    }
  }

  return start;
}

//******************RUN TRANSFER************* */
function runTransfer(options = {}) {
  const auto = !!options.auto;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shIn  = ss.getSheetByName('⏬ ВНЕСЕНИЕ');
  const shProv= ss.getSheetByName('☑️ ПРОВОДКИ');
  const shDict= ss.getSheetByName('Справочник');
  const shActs= ss.getSheetByName('РЕЕСТР АКТОВ');
  const tz    = Session.getScriptTimeZone();
  const BIG_LIMIT = 1e6;

  const rowErrors = [];
  // Счётчики типов ошибок для компактной статистики
  const failureCounts = {
    noWallet: 0,
    noAmount: 0,
    missingArticle: 0,
    missingAct: 0,
    duplicate: 0,
    other: 0
  };

  function err(rowIdx, msg) {
    // Категоризация ошибки по тексту для статистики
    const lower = String(msg).toLowerCase();
    if (lower.includes('кошел')) failureCounts.noWallet++;
    else if (lower.includes('сум') || lower.includes('равна 0')) failureCounts.noAmount++;
    else if (lower.includes('статья') || lower.includes('тип') || lower.includes('катег')) failureCounts.missingArticle++;
    else if (lower.includes('акт')) failureCounts.missingAct++;
    else if (lower.includes('дубл') || lower.includes('повтор')) failureCounts.duplicate++;
    else failureCounts.other++;

    rowErrors.push(`B${10 + rowIdx}: ${msg}`);
  }

  // Очистим возможные нежелательные пробельные символы в B..F перед обработкой
  normalizeInputBF_(shIn);
  const inRange = shIn.getRange('B10:L40');
  const inVals  = inRange.getValues();   // [ [B..L], ... ]

  /* === Проверка месяца дат перед проведением (оставляем как было) === */
  (function precheckMonth_() {
    const now  = new Date();
    const curY = now.getFullYear();
    const curM = now.getMonth(); // 0..11

    const pastIdx   = [];
    const futureIdx = [];

    for (let i = 0; i < inVals.length; i++) {
      const row = inVals[i];
      const amount = row[2]; // D
      const hasAmount = amount !== '' && amount != null;
      if (!hasAmount) continue;

      const d = parseSheetDate_(row[0], Session.getScriptTimeZone());
      if (!d) continue;

      const y = d.getFullYear();
      const m = d.getMonth();

      if (y < curY || (y === curY && m < curM)) pastIdx.push(i);
      else if (y > curY || (y === curY && m > curM)) futureIdx.push(i);
    }

    if (pastIdx.length === 0 && futureIdx.length === 0) return;

    if (pastIdx.length > 0) {
      if (!auto) {
        const btn = confirmDialog_(
          'Проверка дат (прошлый месяц)',
          `Камрад, ты проводишь прошлый месяц (${pastIdx.length} строк). Так и надо?`
        );
        if (!btn) {
          for (const i of pastIdx) {
            const d = parseSheetDate_(inVals[i][0], Session.getScriptTimeZone());
            if (!d) continue;
            inVals[i][0] = adjustDateToCurrentMonthClamp_(d);
          }
        }
      }
    }

    if (futureIdx.length > 0) {
      if (!auto) {
        const btn = confirmDialog_(
          'Проверка дат (будущий месяц)',
          `Камрад, ты проводишь будущий месяц (${futureIdx.length} строк). Так и надо?`
        );
        if (!btn) {
          for (const i of futureIdx) {
            const d = parseSheetDate_(inVals[i][0], Session.getScriptTimeZone());
            if (!d) continue;
            inVals[i][0] = adjustDateToCurrentMonthClamp_(d);
          }
        }
      }
    }

    const dateCol = inVals.map(r => [r[0]]);
    shIn.getRange(10, 2, dateCol.length, 1).setValues(dateCol);
  })();

  /* === Решаем, нужен ли вообще РЕЕСТР АКТОВ в этом запуске === */
  let needActsGrid = false;
  for (let i = 0; i < inVals.length; i++) {
    const row = inVals[i];
    const amount = row[2]; // D
    const hasAmount = amount !== '' && amount != null && Number(amount) !== 0;
    if (!hasAmount) continue;

    const artE   = row[3]; // E
    const altArt = row[6]; // H
    const baseArt = artE || altArt || '';

    if (baseArt === '% Мастер' || baseArt === 'Возврат удержания' || baseArt === 'Выручка по акту') {
      needActsGrid = true;
      break;
    }
  }

  /* === Справочник статей === */
  let dict = [];
  if (shDict.getLastRow() > 1) {
    dict = shDict.getRange(2, 1, shDict.getLastRow() - 1, 5).getValues();
  }

  const pairs  = new Set();   // "статья|расшифровка"
  const acts   = new Map();   // статья → нужен акт
  const hashes = new Set();   // "хэш-статьи" — d начинается с "#"
  const meta   = new Map();   // статья → {t,c,req}
  const byDec  = new Map();   // расшифровка → Set(статей)

  dict.forEach(r => {
    const [t, c, a, d, req] = r;
    if (!a) return;

    pairs.add(a + '|' + d);

    if (String(req).toLowerCase() === 'акт') acts.set(a, true);

    if (String(d).startsWith('#')) hashes.add(a);

    if (!meta.has(a)) meta.set(a, { t, c, req });

    if (d != null && d !== '') {
      const keyDec = String(d).trim();
      if (!byDec.has(keyDec)) byDec.set(keyDec, new Set());
      byDec.get(keyDec).add(a);
    }
  });

  /* === Дубли по последним 100 строкам ПРОВОДОК (оставляем) === */
  const existing   = new Set(); // ключ дубля
  const lastProvRow= shProv.getLastRow();

  if (lastProvRow > 1) {
    const dupWindowSize = 100;
    const startDupRow   = Math.max(2, lastProvRow - dupWindowSize + 1);
    const dupHeight     = lastProvRow - startDupRow + 1;

    const provDup = shProv
      .getRange(startDupRow, 1, dupHeight, 10) // A:J
      .getValues();

    provDup.forEach(r => {
      const [date, wallet, sum, art, dec, act] = r;
      if (date && art && dec && sum !== '' && sum != null) {
        const key = `${fmtDate(date, tz)}|${art}|${dec}|${Number(sum)}`;
        existing.add(key);
      }
    });
  }

  /* === РЕЕСТР АКТОВ (только ключи и флаги, без сумм и остатков) === */
  const ACTS_COL = {
    ADDR: 2,
    ACTNO: 3,
    WAGE_BY_ACT: 9,   // I
    DEPOSIT: 10,      // J
    HANDS: 11,        // K
    MASTER_FLAG: 16,  // P
    RET_FLAG: 17,     // Q
    PAID_FLAG: 18     // R
  };

  let actsGrid = null;
  const keyToRow = {}; // "адрес|акт" → номер строки в РЕЕСТРЕ

  if (needActsGrid && shActs && shActs.getLastRow() > 1) {
    const lastActsRow = shActs.getLastRow();
    actsGrid = shActs.getRange(2, 1, lastActsRow - 1, 18).getValues(); // A:Q

    for (let i = 0; i < actsGrid.length; i++) {
      const row = actsGrid[i];
      const addrCell = row[ACTS_COL.ADDR - 1]; // B
      const actCell  = row[ACTS_COL.ACTNO - 1]; // C
      const key = makeActKey(addrCell, actCell);
      if (!key) continue;
      if (!keyToRow[key]) {
        keyToRow[key] = 2 + i; // реальная строка
      }
    }
  }

  function findActRowByKey_(key) {
    if (!actsGrid) {
      return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'no_data' };
    }
    if (!key) {
      return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'not_found' };
    }

    const row = keyToRow[key];
    if (!row) {
      return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'not_found' };
    }

    const gridIndex = row - 2;
    const gridRow   = actsGrid[gridIndex];

    const paid   = !!gridRow[ACTS_COL.PAID_FLAG   - 1];
    const master = !!gridRow[ACTS_COL.MASTER_FLAG - 1];
    const ret    = !!gridRow[ACTS_COL.RET_FLAG    - 1];

    return { row, gridIndex, paid, master, ret, error: null };
  }

  /* === Сбор результатов === */
  const toWrite        = [];
  const done           = new Set();      // ключи проведённых в этот run
  const toSuggest      = new Map();      // статья → Set(новых расшифровок)
  const DEBUG_REPORT   = false;
  const newDecs        = [];

  const badDate = [], badAct = [], bigDecl = [], dupDecl = [], noDec = [], unknown = [];

  const revenueColorsByRow = {};      // row → color (E-колонка)
  const masterFlagRows     = new Set(); // строки, где поставили флаг ЗП
  const depFlagRows        = new Set(); // строки, где поставили флаг депозита

  const processedRows = new Set();    // индексы строк ⏬ ВНЕСЕНИЕ, которые успешно проведены

  /* === Основной цикл по строкам ⏬ ВНЕСЕНИЕ === */

  // Вспомогательная функция: обрабатывает одну строку по индексу i
  function processRow(i) {
    const r = inVals[i];
    let [date, wallet, sum, artE, dec, act, altArt, cat, type, hint, foreman] = r;

    const hasType    = String(type || '').trim() !== '';
    const hasCat     = String(cat  || '').trim() !== '';
    const hasArtEorH = String(artE || '').trim() !== '' || String(altArt || '').trim() !== '';

    if (!hasType || !hasCat || !hasArtEorH) {
      err(i, 'нет типа (J) или категории (I) или статьи (E/H)');
      return;
    }

    // Если дата пустая — подставляем сегодняшнюю
    if (!date) {
      const today = new Date();
      date = today;
      inVals[i][0] = date;
    }

    // Валидации: кошелёк и сумма
    if (!wallet || String(wallet).trim() === '') {
      err(i, 'нет кошелька (C)');
      return;
    }

    const amount = Number(sum);
    if (sum === '' || sum == null || !isFinite(amount) || amount === 0) {
      err(i, 'нет суммы или она равна 0 (D)');
      return;
    }

    if (!isNaN(amount) && Math.abs(amount) > BIG_LIMIT) {
      bigDecl.push(`${article || ''} ${decoding || ''}`);
    }

    const baseArt = artE || altArt || '';
    let article  = baseArt;
    let decoding = dec;

    // Для статей, которые требуют акт — проверяем наличие акта
    if (acts.get(article) && !act) {
      badAct.push(`${article} ${decoding || ''}`);
      err(i, `Камрад, для статьи "${article}" нужен акт`);
      return;
    }

    const key = `${fmtDate(date, tz)}|${article}|${decoding}|${amount}`;
    const isMasterOrRetention = (article === '% Мастер' || article === 'Возврат удержания');

    const alreadyInProv = existing.has(key);
    const alreadyInRun  = done.has(key);

    const isDuplicate = (!isMasterOrRetention && alreadyInProv) || alreadyInRun;

    if (isDuplicate) {
      if (auto) {
        dupDecl.push(`${article} ${decoding || ''}`);
        return;
      }
      const resp = confirmDialog_(
        'Дубль',
        `Такая проводка уже есть:\n${fmtDate(date, tz)} | ${article} | ${decoding} | ${amount}\nВнести повторно?`
      );
      if (!resp) {
        dupDecl.push(`${article} ${decoding || ''}`);
        return;
      }
    }
    done.add(key);

    if (hashes.has(article) && !decoding) {
      noDec.push(`${article}`);
      return;
    }

    const pairKey = article + '|' + decoding;
    if (!pairs.has(pairKey)) {
      if (hashes.has(article)) {
        // ничего
      } else if (meta.has(article)) {
        if (!toSuggest.has(article)) toSuggest.set(article, new Set());
        toSuggest.get(article).add(decoding);
      } else {
        unknown.push(article);
      }
    }

    // Выручка по акту → собираем цвет подсветки
    if (article === 'Выручка по акту') {
      const actKey = makeActKey(decoding, act);
      const rowActs = actKey ? keyToRow[actKey] : null;
      const color = WALLET_COLORS[wallet] || null;
      if (rowActs && color) {
        revenueColorsByRow[rowActs] = color;
      }
    }

    // Логика по актам для % Мастер / Возврат удержания
    const isMaster    = (article === '% Мастер');
    const isRetention = (article === 'Возврат удержания');

    if (isMaster || isRetention) {
      if (!shActs || !actsGrid) {
        err(i, 'РЕЕСТР АКТОВ не найден или пуст, не могу привязать выплату к акту');
        return;
      }
      if (!decoding || String(decoding).trim() === '') {
        err(i, 'Для "% Мастер"/"Возврат удержания" в F должен быть адрес (как в РЕЕСТР АКТОВ!B)');
        return;
      }
      if (!act || String(act).trim() === '' || String(act).indexOf('АКТ') === -1) {
        err(i, 'В G должен быть номер акта со словом "АКТ" (как в РЕЕСТР АКТОВ!C)');
        return;
      }

      const actKey = makeActKey(decoding, act);
      const res    = findActRowByKey_(actKey);

      if (!res.row) {
        if (res.error === 'not_found') {
          err(i, 'Акт не найден в РЕЕСТР АКТОВ по адресу+акту');
        } else {
          err(i, 'РЕЕСТР АКТОВ не готов (нет данных)');
        }
        return;
      }

      const targetCol   = isMaster ? ACTS_COL.MASTER_FLAG : ACTS_COL.RET_FLAG;
      const alreadyFlag = isMaster ? res.master : res.ret;

      if (alreadyFlag) {
        if (auto) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
        const ask2 = confirmDialog_(
          'Повторная операция по акту',
          'Камрад, по этому акту уже стояла галочка выплаты. Повторить операцию?'
        );
        if (!ask2) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
      }

      actsGrid[res.gridIndex][targetCol - 1] = true;
      if (isMaster) masterFlagRows.add(res.row);
      else          depFlagRows.add(res.row);
    }

    // авто-зеркалирование переводов между кошельками
    const { extraRow, error, required } = handleInternalTransfer_(
      [date, wallet, amount, article, decoding, act, cat, type, hint, foreman]
    );
    if (required && error) {
      err(i, error);
      return;
    }

    // исходная строка
    toWrite.push([date, wallet, amount, article, decoding, act, cat, type, hint, foreman]);
    processedRows.add(i);

    // зеркальная (если есть)
    if (extraRow) {
      toWrite.push(extraRow);
    }
  }

  // Прогоним обработчик по всем строкам
  for (let i = 0; i < inVals.length; i++) {
    const r = inVals[i];
    const isBlankRow = r.every(v => v == null || String(v).trim() === '');
    if (isBlankRow) continue;
    processRow(i);
  }

  /* === Запись в ☑️ ПРОВОДКИ === */
  if (toWrite.length) {
    const curFilter = shProv.getFilter();
    if (curFilter) curFilter.remove();

    const start = findStartRowForProv_(shProv);
    shProv.getRange(start, 1, toWrite.length, 10).setValues(toWrite);
    colorRows_(shProv, start, toWrite);

    const newLast = start + toWrite.length - 1;
    PropertiesService.getDocumentProperties()
      .setProperty('LAST_PROV_ROW', String(newLast));

    // Нативный быстрый Toast больше не показываем здесь — единый итоговый toast будет в конце
  }

  /* === Очистка/сохранение вводимых строк в ⏬ ВНЕСЕНИЕ ===
     — Чистим диапазон B10:G40 (содержимое и форматирование) для проведённых строк
     — Возвращаем только те строки, которые НЕ были проведены (B..G остаются)
  */
  const height = inVals.length;
  const outVals = [];

  // Собираем новые значения для B..G
  for (let i = 0; i < height; i++) {
    const row = inVals[i];
    const isBlankRow = row.every(v => v == null || String(v).trim() === '');
    if (processedRows.has(i) || isBlankRow) {
      outVals.push(['', '', '', '', '', '']);
    } else {
      // возвращаем исходные (или уже поправленные датой) значения B..G
      outVals.push([row[0], row[1], row[2], row[3], row[4], row[5]]);
    }
  }
  // Записываем B..G
  shIn.getRange(IN_START_ROW, IN_COL_B, height, 6).setValues(outVals);

  // Форматирование и заметки не трогаем — очищаем только значения (они уже записаны выше в B..G)
  // (Оставляем форматирование и примечания на месте по просьбе пользователя.)

  /* === Новые расшифровки — как раньше === */
  if (toSuggest.size) {
    // Быстрая нативная проверка: спрашиваем один раз, добавляем по выбору
    const ui = SpreadsheetApp.getUi();
    const wantAddBtn = ui.alert('Новые расшифровки', 'Камрад, я вижу новые расшифровки. Хочешь добавить их в справочник?', ui.ButtonSet.YES_NO);
    if (wantAddBtn === ui.Button.YES) {
      const batchBtn = ui.alert('Режим добавления', 'Добавить все сразу (Да) или по одной с подтверждением (Нет)?', ui.ButtonSet.YES_NO);
      const addAllAtOnce = (batchBtn === ui.Button.YES);

      // Для режима "Добавить все сразу" собираем строки и добавляем батчем
      const rowsToAppend = [];

      toSuggest.forEach((set, art) => {
        if (!meta.has(art)) return;
        const m = meta.get(art);

        const arr = Array.from(set)
          .map(d => (d == null ? '' : String(d).trim()))
          .filter(d => d !== '')
          .filter((d, i, a) => a.indexOf(d) === i)
          .sort((a, b) => a.localeCompare(b, 'ru'));

        if (!arr.length) return;

        if (addAllAtOnce) {
          arr.forEach(d => {
            rowsToAppend.push([m.t, m.c, art, d, m.req]);
            newDecs.push(`${art} — ${d}`);
          });
        } else {
          arr.forEach(d => {
            const resp = ui.alert('Добавить в "Справочник"?', `Тип: ${m.t}\nКатегория: ${m.c}\nСтатья: ${art}\nРасшифровка: ${d}\n\nДобавить эту строку?`, ui.ButtonSet.YES_NO);
            if (resp === ui.Button.YES) {
              shDict.appendRow([m.t, m.c, art, d, m.req]);
              newDecs.push(`${art} — ${d}`);
            }
          });
        }
      });

      if (rowsToAppend.length) {
        const last = shDict.getLastRow();
        const startRow = Math.max(2, last + 1);
        shDict.getRange(startRow, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
      }
    }
  }

  /* === Запись флагов в РЕЕСТР АКТОВ батчем === */
  if (shActs) {
    (function applyActsFlags_() {
      const lastActsRow = shActs.getLastRow();
      if (lastActsRow <= 1) return;
      const height = Math.max(1, lastActsRow - 1);

      function setFlagColumn(colIndex, rowsSet) {
        if (!rowsSet || rowsSet.size === 0) return;
        const colRange = shActs.getRange(2, colIndex, height, 1);
        const colVals = colRange.getValues();
        rowsSet.forEach(r => {
          const idx = r - 2;
          if (idx >= 0 && idx < colVals.length) colVals[idx][0] = true;
        });
        colRange.setValues(colVals);
      }

      setFlagColumn(ACTS_COL.MASTER_FLAG, masterFlagRows);
      setFlagColumn(ACTS_COL.RET_FLAG,    depFlagRows);
    })();

    // 1) Подсветка выручки по акту (E) — батчем по найденным строкам
    (function applyRevenueColors_(){
      const keys = Object.keys(revenueColorsByRow).map(k => Number(k)).filter(n => Number.isFinite(n));
      if (!keys.length) return;
      const minRow = Math.min(...keys);
      const maxRow = Math.max(...keys);
      const height = maxRow - minRow + 1;
      const bg = Array.from({length: height}, () => [null]);
      keys.forEach(r => {
        const color = revenueColorsByRow[String(r)];
        if (color) bg[r - minRow][0] = color;
      });
      shActs.getRange(minRow, 5, height, 1).setBackgrounds(bg);
    })();

    // 2) Полные выплаты ЗП/депозита — зелёный фон + зачёркнутый текст в K / J (применяем блочно)
    function applyStyleBlocks(colIndex, rowsSet) {
      if (!rowsSet || rowsSet.size === 0) return;
      const rows = Array.from(rowsSet).sort((a,b)=>a-b);
      const minRow = rows[0];
      const maxRow = rows[rows.length - 1];
      const height = maxRow - minRow + 1;

      // Забираем существующие значения фонов/цветов/заметок, чтобы не перезаписывать лишнее
      const rng = shActs.getRange(minRow, colIndex, height, 1);
      const existingBG = rng.getBackgrounds();
      const existingFontColors = rng.getFontColors();
      const existingNotes = rng.getNotes();

      // Помечаем нужные строки внутри диапазона
      rows.forEach(r => {
        const idx = r - minRow;
        existingBG[idx][0] = COLOR_BG_FULL_GREEN;
        existingFontColors[idx][0] = COLOR_FONT_DARKGREEN;
        existingNotes[idx][0] = '';
      });

      // Пишем батчем фон/цвет/заметки (меньше вызовов API)
      rng.setBackgrounds(existingBG);
      rng.setFontColors(existingFontColors);
      rng.setNotes(existingNotes);

      // Для зачёркивания (setFontLine) вызываем только для блоков подряд идущих строк
      let blockStart = rows[0];
      let prev = rows[0];
      for (let i = 1; i <= rows.length; i++) {
        const cur = rows[i];
        if (!cur || cur !== prev + 1) {
          const len = prev - blockStart + 1;
          shActs.getRange(blockStart, colIndex, len, 1).setFontLine('line-through');
          blockStart = cur;
        }
        prev = cur;
      }
    }

    applyStyleBlocks(ACTS_COL.HANDS, masterFlagRows);
    applyStyleBlocks(ACTS_COL.DEPOSIT, depFlagRows);
  }

  /* === Финальный отчёт === */
  const lines = [`Перенесено: ${toWrite.length}`];

  if (newDecs.length) {
    lines.push('', 'Добавлены новые расшифровки:');
    newDecs.forEach(r => lines.push('• ' + r));
  }

  if (DEBUG_REPORT) {
    if (badDate.length)  lines.push(`\nБез даты: ${badDate.length}`);
    if (badAct.length)   lines.push(`Без акта: ${badAct.length}`);
    if (bigDecl.length)  lines.push(`Крупные суммы (отклонено): ${bigDecl.length}`);
    if (dupDecl.length)  lines.push(`Дубликаты (отклонено): ${dupDecl.length}`);
    if (noDec.length)    lines.push(`Статьи с # без расшифровки: ${noDec.length}`);
    if (unknown.length)  lines.push(`Неизвестные статьи: ${[...new Set(unknown)].length}`);
  }

  if (rowErrors.length) {
    lines.push('', 'Не проведены (причины):');
    rowErrors.slice(0, 30).forEach(m => lines.push('• ' + m));
    if (rowErrors.length > 30) {
      lines.push(`... и ещё ${rowErrors.length - 30}`);
    }
  }

  // Всегда показываем краткий toast с итогами; подробный отчёт логируем в консоль.
  const summaryParts = [`Перенесено: ${toWrite.length}`];
  if (rowErrors.length) summaryParts.push(`Не проведено: ${rowErrors.length}`);
  if (newDecs.length)    summaryParts.push(`Добавлено расшифровок: ${newDecs.length}`);
  const summary = summaryParts.join('. ');
  SpreadsheetApp.getActive().toast(summary, 'Готово', 8);

  // Логируем подробности и статистику для отладки (можно перенести в отдельный лист при необходимости)
  const stats = [];
  if (failureCounts.noWallet) stats.push(`Нет кошелька: ${failureCounts.noWallet}`);
  if (failureCounts.noAmount) stats.push(`Нет суммы/0: ${failureCounts.noAmount}`);
  if (failureCounts.missingArticle) stats.push(`Нет статьи/категории/типа: ${failureCounts.missingArticle}`);
  if (failureCounts.missingAct) stats.push(`Нет акта: ${failureCounts.missingAct}`);
  if (failureCounts.duplicate) stats.push(`Дубликаты: ${failureCounts.duplicate}`);
  if (failureCounts.other) stats.push(`Прочие ошибки: ${failureCounts.other}`);
  console.info(lines.join('\n'));
  if (stats.length) console.info('Статистика ошибок: ' + stats.join('; '));

  // Если были новые расшифровки — покажем интерактивный диалог добавления (как и раньше)
  // (оставляем существующую логику выше, она уже обработана до этого шага).
}

/* === Coloring === */
function colorRows_(sh, start, rows) {
  const n = rows.length;
  const sumColors = [], walletColors = [];
  rows.forEach(r => {
    const wallet = r[1], type = r[7];
    let cSum = null;
    if (type === 'Доход') cSum = '#E6F4EA';
    if (type === 'Расход') cSum = '#FDEAEA';
    sumColors.push([cSum]);
    let cW = null;
    if (wallet === 'Р/С Строймат') cW = '#DDEBF7';
    else if (wallet === 'Р/С Брендмар') cW = '#FFF2CC';
    else if (wallet === 'Наличные') cW = '#E2EFDA';
    else if (wallet === 'Карта') cW = '#D9F0F0';
    else if (wallet === 'Карта Артема') cW = '#E6E0EC';
    walletColors.push([cW]);
  });
  sh.getRange(start, 3, n, 1).setBackgrounds(sumColors);
  sh.getRange(start, 2, n, 1).setBackgrounds(walletColors);
}

/* === Date helpers — Удалено: setToday/setYesterday/fillDate_ (устаревшие) === */
// Ранее здесь были вспомогательные функции для быстрой установки даты, но
// они удалены как рудименты по запросу владельца проекта.


/* === Internal transfers mirroring === */

// допустимые названия кошельков (для проверки расшифровки)
function allowedWallets_() {
  return new Set([
    'Р/С Строймат',
    'Р/С Брендмар',
    'Наличные',
    'Карта',
    'Карта Артема',
    'Карта Паши',
    'ИП Паши'
  ]);
}

/**
 * Если статья = "Перевод на кошелек" или "Пополнение кошелька",
 * строит вторую (зеркальную) проводку.
 *
 * @param {Array} row [date, wallet, amount, article, decoding, act, category, type, hint, foreman] — как в toWrite
 * @returns {{extraRow: Array|null, error: string|null, required: boolean}}
 *   required=true означает, что для этой строки проверка обязательна (это перевод).
 *   Если error != null — строку проводить нельзя (нет валидной расшифровки-кошелька).
 */
function handleInternalTransfer_(row) {
  const [date, wallet, amount, article, decoding] = row;
  const wallets = allowedWallets_();

  const isOut = article === 'Перевод на кошелек';
  const isIn  = article === 'Пополнение кошелька';

  // не перевод — ничего не делаем
  if (!isOut && !isIn) {
    return { extraRow: null, error: null, required: false };
  }

  // для перевода расшифровка обязательна и должна быть валидным кошельком
  if (!decoding || !wallets.has(decoding)) {
    const msg = isOut
      ? 'Камрад, при "Перевод на кошелек" в расшифровке должен быть целевой кошелёк.'
      : 'Камрад, при "Пополнение кошелька" в расшифровке должен быть исходный кошелёк.';
    return { extraRow: null, error: msg, required: true };
  }

  // строим зеркальную строку
  // дата — та же; сумма — та же
  // кошелёк = расшифровка исходной
  // расшифровка = кошелёк исходной
  // категория = "Перевод м/у счетами"
  // тип = инвертированный
  // статья = "Пополнение кошелька" (для исходного "Перевод на кошелек") ИЛИ наоборот
  const mirrorType     = isOut ? 'Доход'  : 'Расход';
  const mirrorArticle  = isOut ? 'Пополнение кошелька' : 'Перевод на кошелек';
  const mirrorWallet   = decoding;  // куда зачисляем / откуда списываем
  const mirrorDecoding = wallet;    // парный кошелёк для связки

  const extraRow = [
    date,                    // Дата
    mirrorWallet,            // Кошелёк (второй)
    Number(amount),          // Сумма
    mirrorArticle,           // Статья
    mirrorDecoding,          // Расшифровка
    '',                      // Акт (пусто)
    'Перевод м/у счетами',   // Категория (фикс)
    mirrorType,              // Тип (инверсия)
    '',                      // Подсказка
    ''                       // Прораб
  ];

  return { extraRow, error: null, required: true };
}

// fillDate_ удалена — устаревшая функция (setToday/setYesterday удалены)

// normalizeInputBF_ moved to `10-utils.js`
/* === Utils === */

// parseSheetDate_, lastDayOfMonth_, adjustDateToCurrentMonthClamp_, fmtDate, label moved to `10-utils.js`


/** Показывает список статей (кроме "хэш-статей") и возвращает {article, created} либо null */
function pickArticleInteractive_(ui, meta, hashes, dictSheet, byDec, decoding) {
  const articles = Array.from(meta.keys())
    .filter(a => !hashes.has(a))
    .sort((x, y) => String(x).localeCompare(String(y), 'ru'));

  const lines = ['0. [Создать новую статью]']
    .concat(articles.map((a, i) => `${i+1}. ${a}`))
    .join('\n');

  const respData = promptDialog_('К какой статье отнесём эту проводку?', `Расшифровка: ${String(decoding)}\n\nВведи номер:\n\n${lines}`, '');
  if (respData.button !== 'Ok') return null;

  const n = Number(String(respData.text).trim());
  if (Number.isInteger(n) && n >= 1 && n <= articles.length) {
    return { article: articles[n-1], created: false };
  }
  if (n !== 0) return null; // не 0 и не валидный номер → выходим

  // Создание новой статьи
  const nameResp = promptDialog_('Создание статьи', 'Введи название статьи:', '');
  if (nameResp.button !== 'Ok') return null;
  const newArticle = String(nameResp.text).trim();
  if (!newArticle) return null;
  if (meta.has(newArticle)) return { article: newArticle, created: false };

  // Списки типов/категорий из meta
  const types = Array.from(new Set(Array.from(meta.values()).map(m => m.t))).sort((a,b)=>String(a).localeCompare(String(b),'ru'));
  const cats  = Array.from(new Set(Array.from(meta.values()).map(m => m.c))).sort((a,b)=>String(a).localeCompare(String(b),'ru'));

  function chooseFromList_(title, items) {
    const menu = ['0. [Ввести вручную]'].concat(items.map((v,i)=>`${i+1}. ${v}`)).join('\n');
    const r = promptDialog_(title, `Выбери номер:\n\n${menu}`, '');
    if (r.button !== 'Ok') return null;
    const k = Number(String(r.text).trim());
    if (Number.isInteger(k) && k>=1 && k<=items.length) return items[k-1];
    if (k === 0) {
      const r2 = promptDialog_(title, 'Введи значение:', '');
      if (r2.button !== 'Ok') return null;
      const v = String(r2.text).trim();
      return v || null;
    }
    return null;
  }

  const t = chooseFromList_('Выбор типа', types);     if (t == null) return null;
  const c = chooseFromList_('Выбор категории', cats); if (c == null) return null;

  const needAct = confirmDialog_('Требуется акт?', 'Для этой статьи нужен акт?');
  const req = needAct ? 'акт' : '';

  // Запишем новую статью и текущую расшифровку в «Справочник»
  dictSheet.appendRow([t, c, newArticle, String(decoding).trim(), req]);

  // Обновим индексы meta/byDec (pairs добьём в месте вызова)
  meta.set(newArticle, { t, c, req });
  const kDec = String(decoding).trim();
  if (!byDec.has(kDec)) byDec.set(kDec, new Set());
  byDec.get(kDec).add(newArticle);

  return { article: newArticle, created: true };
}

/** Возвращает {article, created} либо null */
function resolveArticleByDec_(ui, dec, meta, hashes, byDec, dictSheet) {
  const keyDec = String(dec).trim();
  const set = byDec.get(keyDec);
  if (set && set.size === 1) {
    return { article: Array.from(set)[0], created: false };
  }
  return pickArticleInteractive_(ui, meta, hashes, dictSheet, byDec, dec);
}

/**
 * Ищет первую полностью пустую строку в блоке B10:G40 на листе "⏬ ВНЕСЕНИЕ".
 * Пустая = все ячейки B..G === '' / null / пробелы.
 * Возвращает номер строки или null, если нет.
 */
function findFirstEmptyRowInInput_(sh) {
  const startRow = IN_START_ROW;
  const height   = IN_HEIGHT;

  const range = sh.getRange(startRow, IN_COL_B, height, IN_COL_F - IN_COL_B + 1); // B..F — считаем строку занятой только по B..F
  const vals  = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const isEmpty = row.every(v => v == null || String(v).trim() === '');
    if (isEmpty) return startRow + i;
  }
  return null;
}

