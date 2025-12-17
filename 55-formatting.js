// 12-formatting.js — форматирование и цветовая разметка

/**
 * Применяет цвета к строкам в листе ПРОВОДКИ
 * @param {Sheet} sh - лист для раскраски
 * @param {number} start - начальная строка
 * @param {Array} rows - массив строк [[date, wallet, amount, ...], ...]
 */
function colorRows_(sh, start, rows) {
  const n = rows.length;
  const sumColors = [], walletColors = [];
  rows.forEach(r => {
    const wallet = r[1], type = r[7];
    let cSum = null;
    if (type === 'Доход') cSum = '#E6F4EA';
    if (type === 'Расход') cSum = '#FDEAEA';
    sumColors.push([cSum]);
    // Используем цвета из WALLET_COLORS константы
    const cW = WALLET_COLORS[wallet] || null;
    walletColors.push([cW]);
  });
  sh.getRange(start, 3, n, 1).setBackgrounds(sumColors);
  sh.getRange(start, 2, n, 1).setBackgrounds(walletColors);
}

/** Допустимые названия кошельков для проверки переводов */
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
 * @param {Array} row [date, wallet, amount, article, decoding, act, category, type, hint, foreman]
 * @returns {{extraRow: Array|null, error: string|null, required: boolean}}
 */
function handleInternalTransfer_(row) {
  const [date, wallet, amount, article, decoding] = row;
  const wallets = allowedWallets_();

  const isOut = article === 'Перевод на кошелек';
  const isIn  = article === 'Пополнение кошелька';

  if (!isOut && !isIn) {
    return { extraRow: null, error: null, required: false };
  }

  // Создаём нормализованный Set для регистронезависимой проверки
  const walletsLowerSet = new Set(
    Array.from(wallets).map(w => String(w).toLowerCase().trim())
  );
  
  const decodingLower = String(decoding || '').toLowerCase().trim();
  
  if (!decoding || !walletsLowerSet.has(decodingLower)) {
    const msg = isOut
      ? 'Камрад, при "Перевод на кошелек" в расшифровке должен быть целевой кошелёк.'
      : 'Камрад, при "Пополнение кошелька" в расшифровке должен быть исходный кошелёк.';
    return { extraRow: null, error: msg, required: true };
  }

  // Находим оригинальное название кошелька (с правильным регистром)
  const mirrorWalletOriginal = Array.from(wallets).find(
    w => String(w).toLowerCase().trim() === decodingLower
  );

  const mirrorType     = isOut ? 'Доход'  : 'Расход';
  const mirrorArticle  = isOut ? 'Пополнение кошелька' : 'Перевод на кошелек';
  const mirrorWallet   = mirrorWalletOriginal || decoding;
  const mirrorDecoding = wallet;

  const extraRow = [
    date,
    mirrorWallet,
    Number(amount),
    mirrorArticle,
    mirrorDecoding,
    '',
    'Перевод м/у счетами',
    mirrorType,
    '',
    ''
  ];

  return { extraRow, error: null, required: true };
}

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
  if (n !== 0) return null;

  const nameResp = promptDialog_('Создание статьи', 'Введи название статьи:', '');
  if (nameResp.button !== 'Ok') return null;
  const newArticle = String(nameResp.text).trim();
  if (!newArticle) return null;
  if (meta.has(newArticle)) return { article: newArticle, created: false };

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

  // Используем централизованную функцию для добавления статьи
  const success = addArticleToDictionary_(dictSheet, newArticle, decoding, t, c, needAct, meta, byDec);
  if (!success) {
    okDialog_('Ошибка', 'Не удалось добавить статью в справочник');
    return null;
  }

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
