// 60-transfer.js — вспомогательные функции для runTransfer
// Пока только одна функция: validateRowBasic, которая проверяет базовую валидность строки
// и нормализует дату/сумму/поля. Эта функция не выполняет сайд-эффектов (не пишет в листы, не показывает диалоги).

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
