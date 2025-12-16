// 10-utils.js — утилиты и вспомогательные функции
// Перенесено из Внесение.js для лучшей модульности.

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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
 * Год сохраняем исходный.
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

/** Нормализует и очищает диапазон B..F (удаляет NBSP и trim) — не трогает формулы */
function normalizeInputBF_(sh) {
  if (!sh) return;
  const startRow = IN_START_ROW;
  const height = IN_HEIGHT;
  const startCol = IN_COL_B;
  const width = IN_COL_F - IN_COL_B + 1;

  const range = sh.getRange(startRow, startCol, height, width);
  const vals = range.getValues();
  const forms = range.getFormulas();
  let changed = false;

  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      // Не трогаем ячейки с формулой
      const formula = forms[r][c];
      if (formula && formula.toString().trim() !== '') continue;

      const v = vals[r][c];
      if (typeof v === 'string') {
        const newV = v.replace(NBSP_RE, ' ').trim();
        if (newV !== v) {
          vals[r][c] = newV === '' ? '' : newV;
          changed = true;
        }
      }
    }
  }

  if (changed) range.setValues(vals);
}
