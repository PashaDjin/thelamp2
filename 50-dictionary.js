// 50-dictionary.js — работа со справочником статей

/**
 * Загружает справочник и строит индексы для быстрого поиска
 * @param {Sheet} shDict - лист "Справочник"
 * @returns {Object} { pairs, acts, hashes, meta, byDec }
 */
function buildDictionaryIndex_(shDict) {
  const pairs  = new Set();   // "статья|расшифровка"
  const acts   = new Map();   // статья → нужен акт
  const hashes = new Set();   // "хэш-статьи" — d начинается с "#"
  const meta   = new Map();   // статья → {t,c,req}
  const byDec  = new Map();   // расшифровка → Set(статей)

  let dict = [];
  if (shDict && shDict.getLastRow() > 1) {
    dict = shDict.getRange(2, 1, shDict.getLastRow() - 1, 5).getValues();
  }

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

  return { pairs, acts, hashes, meta, byDec };
}

/**
 * Добавляет новые расшифровки в справочник (батчем или по одной)
 * @param {Sheet} shDict
 * @param {Map} toSuggest - статья → Set(расшифровок)
 * @param {Map} meta - статья → {t,c,req}
 * @param {boolean} auto - авто-режим (не показывать диалоги)
 * @returns {Array} newDecs - список добавленных расшифровок ["статья — расшифровка", ...]
 */
function addNewDecodings_(shDict, toSuggest, meta, auto) {
  const newDecs = [];
  if (!toSuggest || toSuggest.size === 0) return newDecs;
  if (auto) return newDecs; // в авто-режиме не добавляем

  const ui = SpreadsheetApp.getUi();
  const wantAddBtn = ui.alert('Новые расшифровки', 'Камрад, я вижу новые расшифровки. Хочешь добавить их в справочник?', ui.ButtonSet.YES_NO);
  if (wantAddBtn !== ui.Button.YES) return newDecs;

  const batchBtn = ui.alert('Режим добавления', 'Добавить все сразу (Да) или по одной с подтверждением (Нет)?', ui.ButtonSet.YES_NO);
  const addAllAtOnce = (batchBtn === ui.Button.YES);

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
    try {
      if (!shDict) throw new Error('Лист Справочник не найден');
      shDict.getRange(startRow, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
    } catch (e) {
      console.error('Ошибка записи в Справочник:', e);
      okDialog_('Ошибка', `Не удалось записать расшифровки: ${e.message}`);
      return [];
    }
  }

  return newDecs;
}

/**
 * Добавляет новую статью в справочник
 * @param {Sheet} shDict - лист Справочник
 * @param {String} article - название статьи
 * @param {String} decoding - расшифровка
 * @param {String} type - тип (Доход/Расход)
 * @param {String} category - категория
 * @param {Boolean} needAct - требуется ли акт
 * @param {Map} meta - карта метаданных статей (обновляется)
 * @param {Map} byDec - карта расшифровок (обновляется)
 * @returns {Boolean} - успешность операции
 */
function addArticleToDictionary_(shDict, article, decoding, type, category, needAct, meta, byDec) {
  const req = needAct ? 'акт' : '';
  
  try {
    shDict.appendRow([type, category, article, String(decoding).trim(), req]);
    
    // Обновляем индексы
    meta.set(article, { t: type, c: category, req });
    const kDec = String(decoding).trim();
    if (!byDec.has(kDec)) byDec.set(kDec, new Set());
    byDec.get(kDec).add(article);
    
    return true;
  } catch (e) {
    console.error('Ошибка добавления статьи в справочник:', e);
    return false;
  }
}
