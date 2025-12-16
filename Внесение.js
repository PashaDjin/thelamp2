function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Проводки')
    .addItem('🚀 Провести', 'runTransfer')
    .addSeparator()
    .addItem('📅 Сегодня', 'setToday')
    .addItem('📆 Вчера', 'setYesterday')
    .addSeparator()
    .addItem('Провести выручку по актам', 'createRevenueFromActs')
    .addSeparator()
    .addItem('Провести ЗП', 'createMasterFromActs')
    .addItem('Провести возврат депозитов', 'createDepositReturnFromActs')
    .addToUi();
}
const MOSCOW_TZ = 'Europe/Moscow';
// Цвета для форматирования (подгони HEX под фактические из таблицы)
const COLOR_BG_FULL_GREEN  = '#C6E0B4'; // светло-зелёный фон "закрыто"
const COLOR_FONT_DARKGREEN = '#385723'; // тёмно-зелёный текст
const COLOR_BG_PARTIAL_YELL = '#FFF2CC'; // жёлтый фон "частично"

// Цвета по кошелькам для подсветки E в РЕЕСТРЕ АКТОВ
const WALLET_COLORS = {
  'Р/С Строймат': '#2496dd', // как в colorRows_
  'Р/С Брендмар': '#EABB3D',
  'Наличные':     '#0dac50',
  'Карта':        '#17ddee',
  'Карта Артема': '#E6E0EC',
  'Карта Паши':   '#E6E0EC',
  'ИП Паши':      '#D9D9D9'
};

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
    <div style="font-family:Arial,sans-serif;white-space:pre-wrap;">
      ${escapeHtml_(message)}
    </div>
    ${withInput ? `
      <div style="margin-top:12px;">
        <input id="dlg-input" style="width:100%;box-sizing:border-box;padding:6px;" value="${escapeHtml_(defaultValue)}" />
      </div>
    ` : ''}
    <div style="margin-top:14px;display:flex;gap:8px;justify-content:flex-end;">
      ${buttons.map(b => `<button onclick="submitDialog('${b}')" style="padding:6px 12px;">${escapeHtml_(b)}</button>`).join('')}
    </div>
    <script>
      function submitDialog(btn){
        const v = document.getElementById('dlg-input') ? document.getElementById('dlg-input').value : '';
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
          .setDialogResult('${token}', { button: btn, value: v });
      }
    </script>
  `)
    .setWidth(420)
    .setHeight(withInput ? 240 : 200);

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
    Utilities.sleep(50);
  }

  cache.remove(token);
  return null;
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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

  const lines = items.map(it => `• ${it.addr} — ${it.amount} (${it.actNo})`);
  const ok = confirmDialog_(title, `Камрад, оформить проводки по объектам:\n\n${lines.join('\n')}\n\nПродолжаем?`);
  if (!ok) return;

  // Ищем первую пустую строку во "⏬ ВНЕСЕНИЕ" в блоке B10:G40
  const firstRow = findFirstEmptyRowInInput_(shIn);
  if (!firstRow) {
    okDialog_('Нет места', 'Камрад, во "⏬ ВНЕСЕНИЕ" нет свободных строк в диапазоне B10:G40.');
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

  // Готовим массив значений для B..G
  let values = [];

  if (mode === 'REVENUE') {
    // Для выручки по актам: на каждый акт — две строки (Выручка по акту + НРП 3%)
    items.forEach(it => {
      // 1) основная выручка по акту (как раньше)
      values.push([
        '',              // B: дата — остаётся пустой, ты её ставишь отдельными кнопками
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
    // MASTER / DEPOSIT_RETURN — старая логика, по одной строке на акт
    values = items.map(it => ([
      todayDate,    // B: дата
      '',           // C: кошелёк
      it.amount,    // D: сумма
      article,      // E: статья
      it.addr,      // F: расшифровка (адрес)
      it.actNo      // G: акт
    ]));
  }

  const targetRange = shIn.getRange(firstRow, 2, values.length, 6); // B..G
  targetRange.setValues(values);
  // Формат даты для колонки B
  shIn.getRange(firstRow, 2, values.length, 1).setNumberFormat('dd"."mm"."yyyy');

  let msg = `Создано проводок во "⏬ ВНЕСЕНИЕ": ${values.length}.`;
  if (errors.length) {
    msg += `\n\nПропущено строк: ${errors.length}.\nПервые несколько:\n` +
      errors.slice(0, 5).map(e => '• ' + e).join('\n');
  }

  okDialog_('Готово', msg);
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
function runTransfer() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shIn  = ss.getSheetByName('⏬ ВНЕСЕНИЕ');
  const shProv= ss.getSheetByName('☑️ ПРОВОДКИ');
  const shDict= ss.getSheetByName('Справочник');
  const shActs= ss.getSheetByName('РЕЕСТР АКТОВ');
  const tz    = Session.getScriptTimeZone();
  const BIG_LIMIT = 1e6;

  const rowErrors = [];
  function err(rowIdx, msg) {
    rowErrors.push(`B${10 + rowIdx}: ${msg}`);
  }

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

    if (futureIdx.length > 0) {
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
  for (let i = 0; i < inVals.length; i++) {
    const r = inVals[i];
    let [date, wallet, sum, artE, dec, act, altArt, cat, type, hint, foreman] = r;

    // пустая строка — пропускаем и чистим потом
    const isBlankRow = r.every(v => v == null || String(v).trim() === '');
    if (isBlankRow) continue;

    const hasType    = String(type || '').trim() !== '';
    const hasCat     = String(cat  || '').trim() !== '';
    const hasArtEorH = String(artE || '').trim() !== '' || String(altArt || '').trim() !== '';

    if (!hasType || !hasCat || !hasArtEorH) {
      err(i, 'нет типа (J) или категории (I) или статьи (E/H)');
      continue;
    }

    // Если дата пустая — предлагаем подставить сегодняшнюю
    if (!date) {
      if (confirmDialog_('Нет даты', 'Камрад, дата не указана. Поставить сегодняшнюю и провести?')) {
        const today = new Date();
        date = today;
        inVals[i][0] = date;
      } else {
        badDate.push(label(r, tz));
        err(i, 'Нет даты');
        continue;
      }
    }

    const baseArt = artE || altArt || '';
    let article  = baseArt;
    let decoding = dec;

    if (acts.get(article) && !act) {
      badAct.push(`${article} ${decoding || ''}`);
      err(i, `Камрад, для статьи "${article}" нужен акт`);
      continue;
    }

    const amount = Number(sum);
    if (!isNaN(amount) && Math.abs(amount) > BIG_LIMIT) {
      const resp = confirmDialog_(
        'Проверка суммы',
        `Камрад, сумма ${amount} выглядит подозрительно. Провести?`
      );
      if (!resp) {
        bigDecl.push(`${article} ${decoding || ''}`);
        continue;
      }
    }

    const key = `${fmtDate(date, tz)}|${article}|${decoding}|${amount}`;
    const isMasterOrRetention = (article === '% Мастер' || article === 'Возврат удержания');

    const alreadyInProv = existing.has(key);
    const alreadyInRun  = done.has(key);

    const isDuplicate =
      (!isMasterOrRetention && alreadyInProv) ||
      alreadyInRun;

    if (isDuplicate) {
      const resp = confirmDialog_(
        'Дубль',
        `Такая проводка уже есть:\n${fmtDate(date, tz)} | ${article} | ${decoding} | ${amount}\nВнести повторно?`
      );
      if (!resp) {
        dupDecl.push(`${article} ${decoding || ''}`);
        continue;
      }
    }
    done.add(key);

    if (hashes.has(article) && !decoding) {
      noDec.push(`${article}`);
      continue;
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

    // Выручка по акту → подсветка E в реестре
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
        continue;
      }
      if (!decoding || String(decoding).trim() === '') {
        err(i, 'Для "% Мастер"/"Возврат удержания" в F должен быть адрес (как в РЕЕСТР АКТОВ!B)');
        continue;
      }
      if (!act || String(act).trim() === '' || String(act).indexOf('АКТ') === -1) {
        err(i, 'В G должен быть номер акта со словом "АКТ" (как в РЕЕСТР АКТОВ!C)');
        continue;
      }

      const actKey = makeActKey(decoding, act);
      const res    = findActRowByKey_(actKey);

      if (!res.row) {
        if (res.error === 'not_found') {
          err(i, 'Акт не найден в РЕЕСТР АКТОВ по адресу+акту');
        } else {
          err(i, 'РЕЕСТР АКТОВ не готов (нет данных)');
        }
        continue;
      }

      const targetCol   = isMaster ? ACTS_COL.MASTER_FLAG : ACTS_COL.RET_FLAG;
      const alreadyFlag = isMaster ? res.master : res.ret;

      if (alreadyFlag) {
        const ask2 = confirmDialog_(
          'Повторная операция по акту',
          'Камрад, по этому акту уже стояла галочка выплаты. Повторить операцию?'
        );
        if (!ask2) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          continue;
        }
      }

      // ставим флаг в P или Q
      shActs.getRange(res.row, targetCol).setValue(true);
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
      continue;
    }

    // исходная строка
    toWrite.push([date, wallet, amount, article, decoding, act, cat, type, hint, foreman]);
    processedRows.add(i);

    // зеркальная (если есть)
    if (extraRow) {
      toWrite.push(extraRow);
    }
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
  }

  /* === Очистка/сохранение вводимых строк в ⏬ ВНЕСЕНИЕ ===
     — Чистим весь диапазон B10:G40
     — Возвращаем только те строки, которые НЕ были проведены
  */
  const height = inVals.length;
  const outVals = [];

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
  shIn.getRange(10, 2, height, 6).setValues(outVals);

  /* === Новые расшифровки — как раньше === */
  if (toSuggest.size) {
    const wantAdd = confirmDialog_(
      'Новые расшифровки',
      'Камрад, я вижу новые расшифровки. Хочешь добавить их в справочник?'
    );
    if (wantAdd) {
      const batchOrSingle = confirmDialog_(
        'Режим добавления',
        'Добавить все сразу (Да) или по одной с подтверждением (Нет)?'
      );
      const addAllAtOnce = batchOrSingle;

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
            shDict.appendRow([m.t, m.c, art, d, m.req]);
            newDecs.push(`${art} — ${d}`);
          });
        } else {
          arr.forEach(d => {
            const resp = confirmDialog_(
              'Добавить в "Справочник"?',
              `Тип: ${m.t}\nКатегория: ${m.c}\nСтатья: ${art}\nРасшифровка: ${d}\n\nДобавить эту строку?`
            );
            if (resp) {
              shDict.appendRow([m.t, m.c, art, d, m.req]);
              newDecs.push(`${art} — ${d}`);
            }
          });
        }
      });
    }
  }

  /* === Форматирование РЕЕСТРА АКТОВ по результатам === */
  if (shActs) {
    // 1) Подсветка выручки по акту (E)
    Object.keys(revenueColorsByRow).forEach(rowStr => {
      const row = Number(rowStr);
      const color = revenueColorsByRow[rowStr];
      if (!row || !color) return;
      shActs.getRange(row, 5).setBackground(color); // E
    });

    // 2) Полные выплаты ЗП/депозита — зелёный фон + зачёркнутый текст в K / J
    masterFlagRows.forEach(row => {
      const cell = shActs.getRange(row, ACTS_COL.HANDS); // K
      cell.setBackground(COLOR_BG_FULL_GREEN);
      cell.setFontColor(COLOR_FONT_DARKGREEN);
      cell.setFontLine('line-through');
      cell.setNote('');
    });

    depFlagRows.forEach(row => {
      const cell = shActs.getRange(row, ACTS_COL.DEPOSIT); // J
      cell.setBackground(COLOR_BG_FULL_GREEN);
      cell.setFontColor(COLOR_FONT_DARKGREEN);
      cell.setFontLine('line-through');
      cell.setNote('');
    });
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

  okDialog_('Готово', lines.join('\n'));
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

/* === Date helpers === */
function setToday() { fillDate_(0); }
function setYesterday() { fillDate_(-1); }

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

function fillDate_(offset) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('⏬ ВНЕСЕНИЕ');

  const sumsRange  = sh.getRange('D10:D40'); // читаем суммы
  const datesRange = sh.getRange('B10:B40'); // будем писать даты

  const sums  = sumsRange.getValues();      // [[D8],[D9],...]
  const dates = datesRange.getValues();     // [[B8],[B9],...]

  const d = new Date();
  d.setDate(d.getDate() + offset);
  const f = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy');

  for (let i = 0; i < dates.length; i++) {
    const raw = sums[i][0];                // значение из D
    const hasAmount = raw !== '' && raw != null; // 0 допускаем
    if (hasAmount) dates[i][0] = f;        // ставим дату в B
  }

  datesRange.setValues(dates);
  datesRange.setNumberFormat('dd"."mm"."yyyy');
}

/* === Utils === */

/** Парсит дату из Date | числа (серийная) | строки dd.MM.yyyy. Возвращает Date или null. */
function parseSheetDate_(v, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  if (typeof v === 'number' && isFinite(v)) {
    // Серийная дата Google Sheets: 1899-12-30 как ноль
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = v * 24 * 60 * 60 * 1000;
    const d = new Date(epoch.getTime() + ms);
    return isNaN(d.getTime()) ? null : d;
  }

  const s = String(v || '').trim();
  if (!s) return null;

  // dd.MM.yyyy
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) {
    const dd = Number(m[1]), mm = Number(m[2]) - 1, yy = Number(m[3]);
    const d = new Date(yy, mm, dd);
    if (isNaN(d.getTime())) return null;
    // проверка на реально существующую дату
    if (d.getFullYear() !== yy || d.getMonth() !== mm || d.getDate() !== dd) return null;
    return d;
  }

  // Фолбэк на стандартный парсер (ISO и т.п.)
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

/** Последний день месяца (0..11) указанного года */
function lastDayOfMonth_(year, month0) {
  return new Date(year, month0 + 1, 0).getDate();
}

/**
 * Меняет МЕСЯЦ на текущий, день сохраняет; если дня нет — клампит до конца текущего месяца.
 * Год сохраняем исходный (как просил: 30.09.2025 → 30.10.2025, если сейчас октябрь 2025).
 */
function adjustDateToCurrentMonthClamp_(d) {
  const now = new Date();
  const curM = now.getMonth(); // 0..11
  const y = d.getFullYear();
  const day = d.getDate();
  const maxDay = lastDayOfMonth_(y, curM);
  const newDay = Math.min(day, maxDay);
  return new Date(y, curM, newDay);
}


function fmtDate(d,tz){try{return Utilities.formatDate(new Date(d),tz,'dd.MM.yyyy');}catch(e){return '';} }
function label(r,tz){return `${r[3]||'без статьи'} ${r[4]||''}`;}


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
  const startRow = 10;
  const endRow   = 40;
  const height   = endRow - startRow + 1;

  const range = sh.getRange(startRow, 2, height, 6); // B..G
  const vals  = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const isEmpty = row.every(v => v == null || String(v).trim() === '');
    if (isEmpty) return startRow + i;
  }
  return null;
}

