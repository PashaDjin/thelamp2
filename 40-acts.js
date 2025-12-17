/**
 * ═══════════════════════════════════════════════════════════════════════════
 * 40-acts.js — Работа с листом "РЕЕСТР АКТОВ"
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Этот файл отвечает за работу с реестром актов выполненных работ.
 * 
 * Что такое акт?
 * Это документ, подтверждающий выполненные работы. В нём указаны:
 * - Адрес объекта
 * - Номер акта
 * - Сумма выручки
 * - Зарплата прораба
 * - Депозит (залог)
 * - Флаги: оплачено, выдано прорабу, возвращён депозит
 * 
 * Основные задачи:
 * - Быстрый поиск акта по адресу и номеру
 * - Проставление галочек (флагов) при выплатах
 * - Подсветка оплаченных актов цветом
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ──────────────────────────────────────────────────────────────────────────
// КАРТА КОЛОНОК ЛИСТА "РЕЕСТР АКТОВ"
// ──────────────────────────────────────────────────────────────────────────
// Вместо того чтобы писать "колонка 2" или "B", используем понятные имена
const ACTS_COL = {
  ADDR: 2,          // B — Адрес объекта (например, "ул. Ленина, 15")
  ACTNO: 3,         // C — Номер акта (например, "АКТ-123/2025")
  REVENUE: 5,       // E — Сумма акта (выручка)
  WAGE_BY_ACT: 9,   // I — Зарплата по акту (для прораба)
  DEPOSIT: 10,      // J — Депозит (залог)
  HANDS: 11,        // K — Сумма на руки прорабу
  MASTER_FLAG: 16,  // P — Галочка "Зарплата выплачена"
  RET_FLAG: 17,     // Q — Галочка "Депозит возвращён"
  PAID_FLAG: 18     // R — Галочка "Акт оплачен"
};

/**
 * Строит индекс актов для быстрого поиска
 * 
 * Проблема: если актов сотни, искать каждый раз перебором всех строк — медленно.
 * 
 * Решение: один раз прочитать все акты и создать "карту":
 *   "ул.Ленина|АКТ-123" → строка 15
 *   "пр.Мира|АКТ-456"  → строка 27
 * 
 * Теперь поиск мгновенный: берём ключ, смотрим в карту — получаем номер строки.
 * 
 * @param {Sheet} shActs - Лист "РЕЕСТР АКТОВ"
 * @returns {Object} - Объект с двумя полями:
 *   - actsGrid: массив всех строк актов
 *   - keyToRow: карта "адрес|акт" → номер строки
 */
function buildActsIndex_(shActs) {
  const res = { actsGrid: null, keyToRow: {} };
  
  // Если лист пустой или содержит только заголовок — возвращаем пустой индекс
  if (!shActs || shActs.getLastRow() <= 1) return res;

  const lastActsRow = shActs.getLastRow();
  
  // Читаем все строки актов одним запросом (быстрее, чем по одной)
  // Строки 2 и ниже (первая строка — заголовок)
  // Колонки A:R (18 колонок)
  const actsGrid = shActs.getRange(2, 1, lastActsRow - 1, 18).getValues();
  const keyToRow = {};

  // Проходим по всем строкам и строим карту
  for (let i = 0; i < actsGrid.length; i++) {
    const row = actsGrid[i];
    const addrCell = row[ACTS_COL.ADDR - 1];  // Берём адрес (колонка B)
    const actCell  = row[ACTS_COL.ACTNO - 1]; // Берём номер акта (колонка C)
    
    // Создаём уникальный ключ вида "адрес|акт"
    const key = makeActKey(addrCell, actCell);
    if (!key) continue; // Пропускаем пустые строки
    
    // Запоминаем номер строки (i+2, потому что первая строка данных = строка 2)
    if (!keyToRow[key]) {
      keyToRow[key] = 2 + i;
    }
  }

  res.actsGrid = actsGrid;
  res.keyToRow = keyToRow;
  return res;
}

/**
 * Ищет акт в индексе по ключу адрес|акт
 * 
 * Эта функция находит строку акта и возвращает:
 * - Номер строки в таблице
 * - Индекс в массиве actsGrid
 * - Флаги: оплачено ли, выдано ли, возвращён ли депозит
 * 
 * @param {Array} actsGrid - Массив строк актов (из buildActsIndex_)
 * @param {Object} keyToRow - Карта ключей → строк (из buildActsIndex_)
 * @param {string} key - Ключ вида "адрес|акт"
 * @returns {Object} - Информация об акте:
 *   - row: номер строки в таблице (0 если не найден)
 *   - gridIndex: индекс в массиве actsGrid (-1 если не найден)
 *   - paid: true если акт оплачен
 *   - master: true если зарплата выплачена
 *   - ret: true если депозит возвращён
 *   - error: код ошибки ('no_data', 'not_found') или null
 */
function findActRowByKey_(actsGrid, keyToRow, key) {
  // Если индекс не построен — возвращаем ошибку
  if (!actsGrid) {
    return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'no_data' };
  }
  
  // Если ключ пустой — акт не может быть найден
  if (!key) {
    return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'not_found' };
  }

  // Ищем в карте
  const row = keyToRow[key];
  if (!row) {
    return { row: 0, gridIndex: -1, paid: false, master: false, ret: false, error: 'not_found' };
  }

  // Нашли! Берём данные из массива
  const gridIndex = row - 2;      // Преобразуем номер строки в индекс массива
  const gridRow   = actsGrid[gridIndex];

  // Читаем флаги (галочки)
  const paid   = !!gridRow[ACTS_COL.PAID_FLAG   - 1]; // Оплачено?
  const master = !!gridRow[ACTS_COL.MASTER_FLAG - 1]; // Зарплата выплачена?
  const ret    = !!gridRow[ACTS_COL.RET_FLAG    - 1]; // Депозит возвращён?

  return { row, gridIndex, paid, master, ret, error: null };
}

/**
 * Проставляет галочки в колонках P (зарплата) и Q (депозит)
 * 
 * Когда проводим выплату зарплаты или возврат депозита, нужно поставить
 * галочку в соответствующей колонке РЕЕСТР АКТОВ.
 * 
 * Эта функция делает это батчем (сразу для всех строк), чтобы не дёргать
 * API Google Sheets много раз.
 * 
 * @param {Sheet} shActs - Лист "РЕЕСТР АКТОВ"
 * @param {Set<number>} masterFlagRows - Номера строк, где нужно поставить галочку "Зарплата"
 * @param {Set<number>} depFlagRows - Номера строк, где нужно поставить галочку "Депозит"
 */
function applyActsFlags_(shActs, masterFlagRows, depFlagRows) {
  if (!shActs) return;
  
  const lastActsRow = shActs.getLastRow();
  if (lastActsRow <= 1) return; // Нет данных
  
  const height = Math.max(1, lastActsRow - 1); // Количество строк с данными

  /**
   * Вспомогательная функция: проставляет галочки в одной колонке
   * 
   * @param {number} colIndex - Номер колонки (P=16 или Q=17)
   * @param {Set<number>} rowsSet - Множество номеров строк для проставления галочек
   */
  function setFlagColumn(colIndex, rowsSet) {
    if (!rowsSet || rowsSet.size === 0) return; // Нечего проставлять
    
    // Читаем всю колонку одним запросом
    const colRange = shActs.getRange(2, colIndex, height, 1);
    const colVals = colRange.getValues();
    
    // Проставляем true для нужных строк
    rowsSet.forEach(r => {
      const idx = r - 2; // Преобразуем номер строки в индекс массива
      if (idx >= 0 && idx < colVals.length) {
        colVals[idx][0] = true; // Ставим галочку
      }
    });
    
    // Записываем обратно одним запросом (быстро!)
    colRange.setValues(colVals);
  }

  setFlagColumn(ACTS_COL.MASTER_FLAG, masterFlagRows);
  setFlagColumn(ACTS_COL.RET_FLAG,    depFlagRows);
}

/** Подсветка выручки по акту (E) — батчем */
function applyRevenueColors_(shActs, revenueColorsByRow) {
  if (!shActs) return;
  const keys = Object.keys(revenueColorsByRow || {}).map(k => Number(k)).filter(n => Number.isFinite(n));
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
}

/** Зачёркивание + зелёный фон для выплат ЗП/депозита */
function applyStyleBlocks_(shActs, colIndex, rowsSet) {
  if (!shActs || !rowsSet || rowsSet.size === 0) return;
  const rows = Array.from(rowsSet).sort((a,b)=>a-b);
  const minRow = rows[0];
  const maxRow = rows[rows.length - 1];
  const height = maxRow - minRow + 1;

  const rng = shActs.getRange(minRow, colIndex, height, 1);
  const existingBG = rng.getBackgrounds();
  const existingFontColors = rng.getFontColors();
  const existingNotes = rng.getNotes();

  rows.forEach(r => {
    const idx = r - minRow;
    existingBG[idx][0] = COLOR_BG_FULL_GREEN;
    existingFontColors[idx][0] = COLOR_FONT_DARKGREEN;
    existingNotes[idx][0] = '';
  });

  rng.setBackgrounds(existingBG);
  rng.setFontColors(existingFontColors);
  rng.setNotes(existingNotes);

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
