// 60-transfer.js — вспомогательные функции для runTransfer
// Разбиваем большую функцию runTransfer() на более мелкие, управляемые блоки.

/**
 * validateRowBasic
 * @param {Array} row — массив значений B..L для строки
 * @param {Number} rowIdx — индекс строки относительно B10 (0-based)
 * @returns {Object} — { ok, error, date, wallet, amount, article, decoding, act, type, cat, hint, foreman }
 */
function validateRowBasic(row, rowIdx) {
  let [date, wallet, sum, artE, dec, act, altArt, cat, type, hint, foreman] = row;

  const res = { ok: true, error: null, date, wallet, sum, amount: null, article: null, decoding: dec, act, type, cat, hint, foreman };

  const hasType    = String(type || '').trim() !== '';
  const hasCat     = String(cat  || '').trim() !== '';
  const hasArtEorH = String(artE || '').trim() !== '' || String(altArt || '').trim() !== '';

  if (!hasType || !hasCat || !hasArtEorH) {
    res.ok = false;
    res.error = 'нет типа (J) или категории (I) или статьи (E/H)';
    return res;
  }

  // Если дата пустая — откладываем установку даты на вызывающую сторону, но пометим,
  // что её нужно подставить (caller может записать обратно в inVals).
  if (!date) {
    res.wantsToday = true;
  } else {
    res.date = date;
  }

  // Валидации: кошелёк и сумма
  if (!wallet || String(wallet).trim() === '') {
    res.ok = false;
    res.error = 'нет кошелька (C)';
    return res;
  }

  const amount = Number(sum);
  if (sum === '' || sum == null || !isFinite(amount) || amount === 0) {
    res.ok = false;
    res.error = 'нет суммы или она равна 0 (D)';
    return res;
  }

  res.amount = amount;
  res.article = artE || altArt || '';
  res.decoding = dec;

  return res;
}

/**
 * buildExistingEntriesSet
 * Собирает Set ключей последних dupWindowSize проводок из листа ПРОВОДКИ
 * для определения дубликатов.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shProv — лист ПРОВОДКИ
 * @param {Number} dupWindowSize — размер окна проверки дубликатов
 * @param {String} tz — таймзона для форматирования дат
 * @returns {Set<String>} — множество ключей вида "date|article|decoding|amount"
 */
function buildExistingEntriesSet(shProv, dupWindowSize, tz) {
  const existing = new Set();
  const lastProvRow = shProv.getLastRow();

  if (lastProvRow > 1) {
    const startDupRow = Math.max(2, lastProvRow - dupWindowSize + 1);
    const dupHeight = lastProvRow - startDupRow + 1;
    const provDup = shProv.getRange(startDupRow, 1, dupHeight, 10).getValues();

    provDup.forEach(r => {
      const [date, wallet, sum, art, dec, act] = r;
      if (date && art && dec && sum !== '' && sum != null) {
        const key = `${fmtDate(date, tz)}|${art}|${dec}|${Number(sum)}`;
        existing.add(key);
      }
    });
  }

  return existing;
}

/**
 * needsActsGrid
 * Определяет, нужен ли РЕЕСТР АКТОВ для обработки данного набора строк.
 * @param {Array<Array>} inVals — массив строк из ⏬ ВНЕСЕНИЕ
 * @returns {Boolean} — true, если есть хотя бы одна строка с % Мастер / Возврат удержания / Выручка по акту
 */
function needsActsGrid(inVals) {
  for (let i = 0; i < inVals.length; i++) {
    const row = inVals[i];
    const amount = row[2]; // D
    const hasAmount = amount !== '' && amount != null && Number(amount) !== 0;
    if (!hasAmount) continue;

    const artE = row[3]; // E
    const altArt = row[6]; // H
    const baseArt = artE || altArt || '';

    if (baseArt === '% Мастер' || baseArt === 'Возврат удержания' || baseArt === 'Выручка по акту') {
      return true;
    }
  }
  return false;
}

/**
 * processActsRelatedEntry
 * Обрабатывает записи, связанные с актами (% Мастер, Возврат удержания).
 * Проверяет наличие акта в РЕЕСТР АКТОВ и устанавливает флаги.
 * @param {Object} params — параметры обработки
 * @returns {Object} — { ok, error, targetRow, targetCol, gridIndex }
 */
function processActsRelatedEntry(params) {
  const { article, decoding, act, shActs, actsGrid, keyToRow } = params;
  const isMaster = (article === '% Мастер');
  const isRetention = (article === 'Возврат удержания');

  if (!isMaster && !isRetention) {
    return { ok: true };
  }

  if (!shActs || !actsGrid) {
    return { ok: false, error: 'РЕЕСТР АКТОВ не найден или пуст, не могу привязать выплату к акту' };
  }

  if (!decoding || String(decoding).trim() === '') {
    return { ok: false, error: 'Для "% Мастер"/"Возврат удержания" в F должен быть адрес (как в РЕЕСТР АКТОВ!B)' };
  }

  if (!act || String(act).trim() === '' || String(act).indexOf('АКТ') === -1) {
    return { ok: false, error: 'В G должен быть номер акта со словом "АКТ" (как в РЕЕСТР АКТОВ!C)' };
  }

  const actKey = makeActKey(decoding, act);
  const res = findActRowByKey_(actsGrid, keyToRow, actKey);

  if (!res.row) {
    if (res.error === 'not_found') {
      return { ok: false, error: 'Акт не найден в РЕЕСТР АКТОВ по адресу+акту' };
    } else {
      return { ok: false, error: 'РЕЕСТР АКТОВ не готов (нет данных)' };
    }
  }

  // Проверяем bounds для безопасного доступа к actsGrid
  if (res.gridIndex < 0 || res.gridIndex >= actsGrid.length) {
    return { ok: false, error: `Внутренняя ошибка: некорректный индекс акта (${res.gridIndex})` };
  }

  const targetCol = isMaster ? ACTS_COL.MASTER_FLAG : ACTS_COL.RET_FLAG;
  const alreadyFlag = isMaster ? res.master : res.ret;

  return {
    ok: true,
    targetRow: res.row,
    targetCol: targetCol,
    gridIndex: res.gridIndex,
    alreadyFlag: alreadyFlag
  };
}

/**
 * writeProvodkiToSheet
 * Записывает массив проводок в лист ☑️ ПРОВОДКИ.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shProv — лист ПРОВОДКИ
 * @param {Array<Array>} toWrite — массив строк для записи
 * @param {Number} startRow — начальная строка для записи
 * @returns {Object} — { ok, error, newLastRow }
 */
function writeProvodkiToSheet(shProv, toWrite, startRow) {
  if (!toWrite.length) return { ok: true, newLastRow: 0 };

  try {
    if (!shProv) throw new Error('Лист "☑️ ПРОВОДКИ" не найден');
    shProv.getRange(startRow, 1, toWrite.length, 10).setValues(toWrite);
    colorRows_(shProv, startRow, toWrite);

    const newLastRow = startRow + toWrite.length - 1;
    PropertiesService.getDocumentProperties()
      .setProperty('LAST_PROV_ROW', String(newLastRow));

    return { ok: true, newLastRow: newLastRow };
  } catch (e) {
    console.error('Ошибка записи в ПРОВОДКИ:', e);
    return { ok: false, error: `Не удалось записать проводки: ${e.message}` };
  }
}

/**
 * clearProcessedInputRows
 * Очищает проведённые строки в листе ⏬ ВНЕСЕНИЕ (B..G).
 * Строки, которые НЕ были проведены, остаются нетронутыми.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} shIn — лист ВНЕСЕНИЕ
 * @param {Array<Array>} inVals — исходные значения из ВНЕСЕНИЕ
 * @param {Set<Number>} processedRows — индексы успешно проведённых строк
 */
function clearProcessedInputRows(shIn, inVals, processedRows) {
  const height = inVals.length;
  const outVals = [];

  for (let i = 0; i < height; i++) {
    const row = inVals[i];
    const isBlankRow = row.every(v => v == null || String(v).trim() === '');
    if (processedRows.has(i) || isBlankRow) {
      outVals.push(['', '', '', '', '', '']);
    } else {
      outVals.push([row[0], row[1], row[2], row[3], row[4], row[5]]);
    }
  }

  shIn.getRange(IN_START_ROW, IN_COL_B, height, 6).setValues(outVals);
}
