/**
 * ═══════════════════════════════════════════════════════════════════════════
 * 10-utils.js — Утилиты и вспомогательные функции
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Этот файл содержит "чистые функции" — функции, которые:
 * - Не изменяют данные в таблице
 * - Не показывают диалоги
 * - Просто берут входные данные и возвращают результат
 * 
 * Такие функции легко тестировать и переиспользовать в разных местах.
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Экранирует HTML-символы для безопасного отображения в диалогах
 * 
 * Что это значит?
 * Если пользователь введёт "<script>alert('hack')</script>" в расшифровку,
 * эта функция превратит это в безопасный текст, который просто отобразится,
 * а не выполнится как код.
 * 
 * @param {string} s - Строка для экранирования
 * @returns {string} - Безопасная строка для вставки в HTML
 */
function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')   // & → &amp;
    .replace(/</g, '&lt;')    // < → &lt;
    .replace(/>/g, '&gt;')    // > → &gt;
    .replace(/"/g, '&quot;')  // " → &quot;
    .replace(/'/g, '&#39;');  // ' → &#39;
}

/**
 * Парсит (преобразует) дату из разных форматов в объект Date
 * 
 * Google Sheets может хранить даты в трёх форматах:
 * 1. Объект Date — уже готовая дата
 * 2. Число — "серийная дата" (количество дней с 30 декабря 1899)
 * 3. Строка — текст вида "17.12.2025"
 * 
 * Эта функция понимает все три формата и возвращает нормальную дату.
 * 
 * @param {Date|number|string} v - Значение из ячейки таблицы
 * @returns {Date|null} - Объект Date или null, если дату распознать не удалось
 */
function parseSheetDate_(v) {
  // Если это уже Date и дата корректная — возвращаем как есть
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  // Если это число (серийная дата из Google Sheets)
  if (typeof v === 'number' && isFinite(v)) {
    // Google Sheets считает даты от 30 декабря 1899 года
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = v * 24 * 60 * 60 * 1000; // Переводим дни в миллисекунды
    const d = new Date(epoch.getTime() + ms);
    return isNaN(d.getTime()) ? null : d;
  }

  const s = String(v || '').trim();
  if (!s) return null; // Пустая строка — не дата

  // Пытаемся распарсить формат dd.MM.yyyy (например, 17.12.2025)
  const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) {
    const dd = Number(m[1]);      // День
    const mm = Number(m[2]) - 1;  // Месяц (в JavaScript месяцы с 0, поэтому -1)
    const yy = Number(m[3]);      // Год
    const d = new Date(yy, mm, dd);
    
    if (isNaN(d.getTime())) return null; // Дата некорректна
    
    // Проверяем, что дата реально существует (например, 31.02.2025 не существует)
    if (d.getFullYear() !== yy || d.getMonth() !== mm || d.getDate() !== dd) {
      return null;
    }
    return d;
  }

  // Если не получилось — пробуем стандартный парсер JavaScript (ISO формат и др.)
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

/**
 * Возвращает последний день указанного месяца
 * 
 * Например:
 * - Февраль 2025 → 28 (не високосный)
 * - Февраль 2024 → 29 (високосный)
 * - Декабрь → всегда 31
 * 
 * @param {number} year - Год (например, 2025)
 * @param {number} month0 - Месяц от 0 до 11 (0=январь, 11=декабрь)
 * @returns {number} - Последний день месяца (28, 29, 30 или 31)
 */
function lastDayOfMonth_(year, month0) {
  // Трюк: создаём дату первого дня следующего месяца, затем отнимаем 1 день
  return new Date(year, month0 + 1, 0).getDate();
}

/**
 * Корректирует дату на текущий месяц, сохраняя день
 * 
 * Используется когда пользователь вводит дату из прошлого/будущего месяца,
 * но хочет провести операцию в текущем месяце.
 * 
 * Примеры:
 * - 5 ноября → 5 декабря (если сейчас декабрь)
 * - 31 января → 28 февраля (если в феврале нет 31 дня)
 * 
 * @param {Date} d - Исходная дата
 * @returns {Date} - Дата с текущим месяцем
 */
function adjustDateToCurrentMonthClamp_(d) {
  const now = new Date();
  const curM = now.getMonth();        // Текущий месяц (0-11)
  const y = d.getFullYear();          // Год из исходной даты
  const day = d.getDate();            // День из исходной даты
  
  // Проверяем, есть ли такой день в текущем месяце
  const maxDay = lastDayOfMonth_(y, curM);
  const newDay = Math.min(day, maxDay); // Если день больше max — берём последний день месяца
  
  return new Date(y, curM, newDay);
}

/**
 * Форматирует дату в строку вида "17.12.2025"
 * 
 * @param {Date} d - Дата для форматирования
 * @param {string} tz - Таймзона (например, 'Europe/Moscow')
 * @returns {string} - Отформатированная строка или пустая строка при ошибке
 */
function fmtDate(d, tz) {
  try {
    return Utilities.formatDate(new Date(d), tz, 'dd.MM.yyyy');
  } catch(e) {
    return '';
  }
}

/**
 * Создаёт метку для строки проводки (для отладки и логов)
 * 
 * @param {Array} r - Массив значений строки [date, wallet, amount, article, decoding, ...]
 * @param {string} tz - Таймзона
 * @returns {string} - Строка вида "Зарплата Сидорову"
 */
function label(r, tz) {
  return `${r[3] || 'без статьи'} ${r[4] || ''}`;
}

/**
 * Очищает данные в диапазоне B:F от неразрывных пробелов
 * 
 * Проблема: когда копируешь данные из Word/PDF, могут вставиться "невидимые"
 * неразрывные пробелы (NBSP). Они выглядят как обычные, но ломают сравнение строк.
 * 
 * Эта функция находит такие пробелы и заменяет на обычные, затем удаляет
 * пробелы в начале и конце строк (trim).
 * 
 * Важно: не трогает ячейки с формулами!
 * 
 * @param {Sheet} sh - Лист для обработки
 */
function normalizeInputBF_(sh) {
  if (!sh) return;
  
  const startRow = IN_START_ROW;  // Начальная строка (10)
  const height = IN_HEIGHT;        // Количество строк (31)
  const startCol = IN_COL_B;       // Начальная колонка B (2)
  const width = IN_COL_F - IN_COL_B + 1; // Ширина: от B до F (5 колонок)

  const range = sh.getRange(startRow, startCol, height, width);
  const vals = range.getValues();     // Читаем значения
  const forms = range.getFormulas();  // Читаем формулы
  let changed = false;

  // Проходим по всем ячейкам
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      // Пропускаем ячейки с формулами (их нельзя менять)
      const formula = forms[r][c];
      if (formula && formula.toString().trim() !== '') continue;

      const v = vals[r][c];
      if (typeof v === 'string') {
        // Заменяем неразрывные пробелы на обычные и обрезаем лишние пробелы
        const newV = v.replace(NBSP_RE, ' ').trim();
        if (newV !== v) {
          vals[r][c] = newV === '' ? '' : newV;
          changed = true;
        }
      }
    }
  }

  // Записываем обратно только если что-то изменилось
  if (changed) range.setValues(vals);
}
