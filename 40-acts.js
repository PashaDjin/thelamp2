// 40-acts.js — работа с РЕЕСТР АКТОВ и стилями

// Глобальная карта колонок для листа актов
const ACTS_COL = {
  ADDR: 2,          // B
  ACTNO: 3,         // C
  REVENUE: 5,       // E (сумма акта/выручка)
  WAGE_BY_ACT: 9,   // I
  DEPOSIT: 10,      // J
  HANDS: 11,        // K
  MASTER_FLAG: 16,  // P
  RET_FLAG: 17,     // Q
  PAID_FLAG: 18     // R
};

/** Строит actsGrid и keyToRow (адрес|акт → строка) */
function buildActsIndex_(shActs) {
  const res = { actsGrid: null, keyToRow: {} };
  if (!shActs || shActs.getLastRow() <= 1) return res;

  const lastActsRow = shActs.getLastRow();
  const actsGrid = shActs.getRange(2, 1, lastActsRow - 1, 18).getValues(); // A:Q
  const keyToRow = {};

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

  res.actsGrid = actsGrid;
  res.keyToRow = keyToRow;
  return res;
}

/** Ищет строку акта по ключу адрес|акт */
function findActRowByKey_(actsGrid, keyToRow, key) {
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

/** Батчево проставляет флаги MASTER/RET в РЕЕСТР АКТОВ */
function applyActsFlags_(shActs, masterFlagRows, depFlagRows) {
  if (!shActs) return;
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
