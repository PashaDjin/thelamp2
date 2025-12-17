// 70-createFromActs.js — создание проводок из РЕЕСТР АКТОВ

// Экспортируемые функции для меню
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
 *  - mode='REVENUE'          → "Выручка по акту", сумма из E
 *
 * Работает только, если активен лист "РЕЕСТР АКТОВ" и есть выделение.
 */
function createEntriesFromSelectedActs_({ mode }) {
  const ss   = SpreadsheetApp.getActive();
  const shActs = ss.getSheetByName(SHT_ACTS);
  const shIn   = ss.getSheetByName(SHT_IN);

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
      if (row > 1) rowSet.add(row);
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
      mode === 'MASTER'         ? 11 : // K — "на руки"
      mode === 'DEPOSIT_RETURN' ? 10  : // J — возврат депозита
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

  // Нормализуем вход (очистка NBSP + trim) в B..F
  normalizeInputBF_(shIn);

  // Ищем первую пустую строку во "⏬ ВНЕСЕНИЕ" в блоке B10:F40
  const firstRow = findFirstEmptyRowInInput_(shIn);
  if (!firstRow) {
    // Диагностика
    const diagRange = shIn.getRange(IN_START_ROW, IN_COL_B, IN_HEIGHT, IN_COL_F - IN_COL_B + 1);
    const diagVals  = diagRange.getValues();
    const nonEmptyRows = [];
    for (let ri = 0; ri < diagVals.length; ri++) {
      const row = diagVals[ri];
      const cols = [];
      for (let ci = 0; ci < row.length; ci++) {
        const v = row[ci];
        if (v != null && String(v).trim() !== '') {
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

  // REVENUE на каждый акт будет 2 строки (Выручка + НРП)
  const rowsPerItem = (mode === 'REVENUE') ? 2 : 1;
  const lastRowNeeded = firstRow + rowsPerItem * items.length - 1;
  if (lastRowNeeded > 40) {
    okDialog_('Нет места', 'Камрад, не хватает свободных строк во "⏬ ВНЕСЕНИЕ" для всех проводок. Освободи место и попробуй ещё раз.');
    return;
  }

  // Дата по Москве
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
      values.push([
        todayDate,
        '',
        it.amount,
        article,
        it.addr,
        it.actNo
      ]);

      const nrpAmount = Math.round(it.amount * 0.03 * 100) / 100;

      values.push([
        todayDate,
        '',
        nrpAmount,
        'НРП',
        it.addr,
        it.actNo
      ]);
    });
  } else {
    // MASTER / DEPOSIT_RETURN — одна строка на акт
    values = items.map(it => ([
      todayDate,
      '',
      it.amount,
      article,
      it.addr,
      it.actNo
    ]));
  }

  // Перед записью уверимся, что все строки имеют ровно 6 колонок (B..G)
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
  shIn.getRange(firstRow, 2, values.length, 1).setNumberFormat('dd"."mm"."yyyy');

  let msg = `Создано проводок во "⏬ ВНЕСЕНИЕ": ${values.length}.`;
  if (errors.length) {
    msg += `\n\nПропущено строк: ${errors.length}.\nПервые несколько:\n` +
      errors.slice(0, 5).map(e => '• ' + e).join('\n');
  }

  SpreadsheetApp.getActive().toast(msg, 'Готово', 5);
}

/**
 * Ищет первую полностью пустую строку в блоке B10:G40 на листе "⏬ ВНЕСЕНИЕ".
 * Пустая = все ячейки B..F === '' / null / пробелы.
 * Возвращает номер строки или null.
 */
function findFirstEmptyRowInInput_(sh) {
  const startRow = IN_START_ROW;
  const height   = IN_HEIGHT;

  const range = sh.getRange(startRow, IN_COL_B, height, IN_COL_F - IN_COL_B + 1); // B..F
  const vals  = range.getValues();

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const isEmpty = row.every(v => v == null || String(v).trim() === '');
    if (isEmpty) return startRow + i;
  }
  return null;
}
