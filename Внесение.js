function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Проводки')
    .addItem('🚀 Провести', 'runTransfer')
    .addSeparator()

    .addItem('Провести выручку по актам', 'createRevenueFromActs')
    .addSeparator()
    .addItem('Провести ЗП', 'createMasterFromActs')
    .addItem('Провести возврат депозитов', 'createDepositReturnFromActs')
    .addToUi();
}
// Константы перемещены в `00-constants.js`
// Утилиты перемещены в `10-utils.js`


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
    <div style="white-space:pre-wrap;">${escapeHtml_(message)}</div>
    ${withInput ? `<div><input id="dlg-input" value="${escapeHtml_(defaultValue)}" /></div>` : ''}
    <div>${buttons.map(b => `<button onclick="submitDialog('${b}')">${escapeHtml_(b)}</button>`).join('')}</div>
    <script>
      function submitDialog(btn){
        const v = document.getElementById('dlg-input') ? document.getElementById('dlg-input').value : '';
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
          .setDialogResult('${token}', { button: btn, value: v });
      }
      document.addEventListener('DOMContentLoaded', function(){
        const b = document.querySelector('button'); if(b) b.focus();
      });
    </script>
  `)
    .setWidth(380)
    .setHeight(withInput ? 180 : 140);

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
    Utilities.sleep(30);
  }

  cache.remove(token);
  return null;
}

// escapeHtml_ moved to `10-utils.js`
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

// createMasterFromActs, createDepositReturnFromActs, createRevenueFromActs,
// createEntriesFromSelectedActs_, findFirstEmptyRowInInput_ moved to 70-createFromActs.js



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
function runTransfer(options = {}) {
  const auto = !!options.auto;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shIn  = ss.getSheetByName('⏬ ВНЕСЕНИЕ');
  const shProv= ss.getSheetByName('☑️ ПРОВОДКИ');
  const shDict= ss.getSheetByName('Справочник');
  const shActs= ss.getSheetByName('РЕЕСТР АКТОВ');
  const tz    = Session.getScriptTimeZone();
  const BIG_LIMIT = 1e6;

  const rowErrors = [];
  // Счётчики типов ошибок для компактной статистики
  const failureCounts = {
    noWallet: 0,
    noAmount: 0,
    missingArticle: 0,
    missingAct: 0,
    duplicate: 0,
    other: 0
  };

  function err(rowIdx, msg) {
    // Категоризация ошибки по тексту для статистики
    const lower = String(msg).toLowerCase();
    if (lower.includes('кошел')) failureCounts.noWallet++;
    else if (lower.includes('сум') || lower.includes('равна 0')) failureCounts.noAmount++;
    else if (lower.includes('статья') || lower.includes('тип') || lower.includes('катег')) failureCounts.missingArticle++;
    else if (lower.includes('акт')) failureCounts.missingAct++;
    else if (lower.includes('дубл') || lower.includes('повтор')) failureCounts.duplicate++;
    else failureCounts.other++;

    rowErrors.push(`B${10 + rowIdx}: ${msg}`);
  }

  // Очистим возможные нежелательные пробельные символы в B..F перед обработкой
  normalizeInputBF_(shIn);
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
      if (!auto) {
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
    }

    if (futureIdx.length > 0) {
      if (!auto) {
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
  const dictIdx = buildDictionaryIndex_(shDict);
  const pairs  = dictIdx.pairs;
  const acts   = dictIdx.acts;
  const hashes = dictIdx.hashes;
  const meta   = dictIdx.meta;
  const byDec  = dictIdx.byDec;

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
  let actsGrid = null;
  let keyToRow = {};

  if (needActsGrid && shActs && shActs.getLastRow() > 1) {
    const actsIdx = buildActsIndex_(shActs);
    actsGrid = actsIdx.actsGrid;
    keyToRow = actsIdx.keyToRow;
  }

  /* === Сбор результатов === */
  const toWrite        = [];
  const done           = new Set();      // ключи проведённых в этот run
  const toSuggest      = new Map();      // статья → Set(новых расшифровок)
  const DEBUG_REPORT   = false;

  const badDate = [], badAct = [], bigDecl = [], dupDecl = [], noDec = [], unknown = [];

  const revenueColorsByRow = {};      // row → color (E-колонка)
  const masterFlagRows     = new Set(); // строки, где поставили флаг ЗП
  const depFlagRows        = new Set(); // строки, где поставили флаг депозита

  const processedRows = new Set();    // индексы строк ⏬ ВНЕСЕНИЕ, которые успешно проведены

  /* === Основной цикл по строкам ⏬ ВНЕСЕНИЕ === */

  // Вспомогательная функция: обрабатывает одну строку по индексу i
  function processRow(i) {
    const r = inVals[i];
    const basic = validateRowBasic(r, i);
    if (!basic.ok) {
      err(i, basic.error);
      return;
    }

    let { date, wallet, amount, article, decoding, act, cat, type, hint, foreman } = basic;

    // Если дата была пустой — заполняем сегодня и фиксируем в inVals
    if (basic.wantsToday) {
      const today = new Date();
      date = today;
      inVals[i][0] = date;
    }

    if (!isNaN(amount) && Math.abs(amount) > BIG_LIMIT) {
      bigDecl.push(`${article || ''} ${decoding || ''}`);
    }

    // Для статей, которые требуют акт — проверяем наличие акта
    if (acts.get(article) && !act) {
      badAct.push(`${article} ${decoding || ''}`);
      err(i, `Камрад, для статьи "${article}" нужен акт`);
      return;
    }

    const key = `${fmtDate(date, tz)}|${article}|${decoding}|${amount}`;
    const isMasterOrRetention = (article === '% Мастер' || article === 'Возврат удержания');

    const alreadyInProv = existing.has(key);
    const alreadyInRun  = done.has(key);

    const isDuplicate = (!isMasterOrRetention && alreadyInProv) || alreadyInRun;

    if (isDuplicate) {
      if (auto) {
        dupDecl.push(`${article} ${decoding || ''}`);
        return;
      }
      const resp = confirmDialog_(
        'Дубль',
        `Такая проводка уже есть:\n${fmtDate(date, tz)} | ${article} | ${decoding} | ${amount}\nВнести повторно?`
      );
      if (!resp) {
        dupDecl.push(`${article} ${decoding || ''}`);
        return;
      }
    }
    done.add(key);

    if (hashes.has(article) && !decoding) {
      noDec.push(`${article}`);
      return;
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

    // Выручка по акту → собираем цвет подсветки
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
        return;
      }
      if (!decoding || String(decoding).trim() === '') {
        err(i, 'Для "% Мастер"/"Возврат удержания" в F должен быть адрес (как в РЕЕСТР АКТОВ!B)');
        return;
      }
      if (!act || String(act).trim() === '' || String(act).indexOf('АКТ') === -1) {
        err(i, 'В G должен быть номер акта со словом "АКТ" (как в РЕЕСТР АКТОВ!C)');
        return;
      }

      const actKey = makeActKey(decoding, act);
      const res    = findActRowByKey_(actsGrid, keyToRow, actKey);

      if (!res.row) {
        if (res.error === 'not_found') {
          err(i, 'Акт не найден в РЕЕСТР АКТОВ по адресу+акту');
        } else {
          err(i, 'РЕЕСТР АКТОВ не готов (нет данных)');
        }
        return;
      }

      const targetCol   = isMaster ? ACTS_COL.MASTER_FLAG : ACTS_COL.RET_FLAG;
      const alreadyFlag = isMaster ? res.master : res.ret;

      if (alreadyFlag) {
        if (auto) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
        const ask2 = confirmDialog_(
          'Повторная операция по акту',
          'Камрад, по этому акту уже стояла галочка выплаты. Повторить операцию?'
        );
        if (!ask2) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
      }

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
      return;
    }

    // исходная строка
    toWrite.push([date, wallet, amount, article, decoding, act, cat, type, hint, foreman]);
    processedRows.add(i);

    // зеркальная (если есть)
    if (extraRow) {
      toWrite.push(extraRow);
    }
  }

  // Прогоним обработчик по всем строкам
  for (let i = 0; i < inVals.length; i++) {
    const r = inVals[i];
    const isBlankRow = r.every(v => v == null || String(v).trim() === '');
    if (isBlankRow) continue;
    processRow(i);
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

    // Нативный быстрый Toast больше не показываем здесь — единый итоговый toast будет в конце
  }

  /* === Очистка/сохранение вводимых строк в ⏬ ВНЕСЕНИЕ ===
     — Чистим диапазон B10:G40 (содержимое и форматирование) для проведённых строк
     — Возвращаем только те строки, которые НЕ были проведены (B..G остаются)
  */
  const height = inVals.length;
  const outVals = [];

  // Собираем новые значения для B..G
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
  // Записываем B..G
  shIn.getRange(IN_START_ROW, IN_COL_B, height, 6).setValues(outVals);

  // Форматирование и заметки не трогаем — очищаем только значения (они уже записаны выше в B..G)
  // (Оставляем форматирование и примечания на месте по просьбе пользователя.)

  /* === Новые расшифровки === */
  const newDecs = addNewDecodings_(shDict, toSuggest, meta, auto);

  /* === Запись флагов и стилей в РЕЕСТР АКТОВ === */
  if (shActs) {
    applyActsFlags_(shActs, masterFlagRows, depFlagRows);
    applyRevenueColors_(shActs, revenueColorsByRow);
    applyStyleBlocks_(shActs, ACTS_COL.HANDS, masterFlagRows);
    applyStyleBlocks_(shActs, ACTS_COL.DEPOSIT, depFlagRows);
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

  // Всегда показываем краткий toast с итогами; подробный отчёт логируем в консоль.
  const summaryParts = [`Перенесено: ${toWrite.length}`];
  if (rowErrors.length) summaryParts.push(`Не проведено: ${rowErrors.length}`);
  if (newDecs.length)    summaryParts.push(`Добавлено расшифровок: ${newDecs.length}`);
  const summary = summaryParts.join('. ');
  SpreadsheetApp.getActive().toast(summary, 'Готово', 8);

  // Логируем подробности и статистику для отладки (можно перенести в отдельный лист при необходимости)
  const stats = [];
  if (failureCounts.noWallet) stats.push(`Нет кошелька: ${failureCounts.noWallet}`);
  if (failureCounts.noAmount) stats.push(`Нет суммы/0: ${failureCounts.noAmount}`);
  if (failureCounts.missingArticle) stats.push(`Нет статьи/категории/типа: ${failureCounts.missingArticle}`);
  if (failureCounts.missingAct) stats.push(`Нет акта: ${failureCounts.missingAct}`);
  if (failureCounts.duplicate) stats.push(`Дубликаты: ${failureCounts.duplicate}`);
  if (failureCounts.other) stats.push(`Прочие ошибки: ${failureCounts.other}`);
  console.info(lines.join('\n'));
  if (stats.length) console.info('Статистика ошибок: ' + stats.join('; '));

  // Если были новые расшифровки — покажем интерактивный диалог добавления (как и раньше)
  // (оставляем существующую логику выше, она уже обработана до этого шага).
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

/* === Date helpers — Удалено: setToday/setYesterday/fillDate_ (устаревшие) === */
// Ранее здесь были вспомогательные функции для быстрой установки даты, но
// они удалены как рудименты по запросу владельца проекта.


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

// fillDate_ удалена — устаревшая функция (setToday/setYesterday удалены)

// normalizeInputBF_ moved to `10-utils.js`
/* === Utils === */

// parseSheetDate_, lastDayOfMonth_, adjustDateToCurrentMonthClamp_, fmtDate, label moved to `10-utils.js`


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
// findFirstEmptyRowInInput_ moved to 70-createFromActs.js

