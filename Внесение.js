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

// === Все функции перенесены в модули ===
// Константы → 00-constants.js
// Утилиты → 10-utils.js
// UI-диалоги → 20-ui-dialogs.js (showDialogAndWait_, setDialogResult, confirmDialog_, okDialog_, promptDialog_)
// Работа с актами → 40-acts.js
// Справочник → 50-dictionary.js
// Форматирование → 55-formatting.js (colorRows_, handleInternalTransfer_, pickArticleInteractive_, resolveArticleByDec_)
// Валидация → 60-transfer.js
// Создание из актов → 70-createFromActs.js

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
  const startTime = new Date().getTime(); // 🔍 ПРОФИЛИРОВАНИЕ
  const auto = !!options.auto;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const shIn  = ss.getSheetByName('⏬ ВНЕСЕНИЕ');
  const shProv= ss.getSheetByName('☑️ ПРОВОДКИ');
  const shDict= ss.getSheetByName('Справочник');
  const shActs= ss.getSheetByName('РЕЕСТР АКТОВ');
  const tz    = Session.getScriptTimeZone();
  const BIG_LIMIT = 1e6;
  
  function logTime(label) {
    const elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(2);
    console.log(`⏱️ [${elapsed}s] ${label}`);
  }

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
  logTime('normalizeInputBF_ завершена');
  
  const inRange = shIn.getRange('B10:L40');
  const inVals  = inRange.getValues();   // [ [B..L], ... ]
  logTime('чтение inVals завершено');

  /* === Проверка месяца дат перед проведением (оптимизировано) === */
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

      const d = parseSheetDate_(row[0]);
      if (!d) {
        console.warn(`Строка B${10 + i}: некорректная дата "${row[0]}", пропускаем`);
        continue;
      }

      const y = d.getFullYear();
      const m = d.getMonth();

      if (y < curY || (y === curY && m < curM)) pastIdx.push(i);
      else if (y > curY || (y === curY && m > curM)) futureIdx.push(i);
    }

    if (pastIdx.length === 0 && futureIdx.length === 0) return;

    // 🚀 ОПТИМИЗАЦИЯ: объединяем диалоги в один
    if (!auto && (pastIdx.length > 0 || futureIdx.length > 0)) {
      let msg = '';
      if (pastIdx.length > 0) {
        msg += `Прошлый месяц: ${pastIdx.length} строк\n`;
      }
      if (futureIdx.length > 0) {
        msg += `Будущий месяц: ${futureIdx.length} строк\n`;
      }
      msg += '\nИсправить даты на текущий месяц?';

      const btn = confirmDialog_('Проверка дат', msg);
      
      if (!btn) { // Пользователь выбрал "Нет" — исправляем даты
        for (const i of pastIdx) {
          const d = parseSheetDate_(inVals[i][0]);
          if (!d) continue;
          inVals[i][0] = adjustDateToCurrentMonthClamp_(d);
        }
        for (const i of futureIdx) {
          const d = parseSheetDate_(inVals[i][0]);
          if (!d) continue;
          inVals[i][0] = adjustDateToCurrentMonthClamp_(d);
        }
      }
    }

    // 🚀 ОПТИМИЗАЦИЯ: записываем даты один раз
    const dateCol = inVals.map(r => [r[0]]);
    shIn.getRange(10, 2, dateCol.length, 1).setValues(dateCol);
  })();
  logTime('precheckMonth_ завершена');

  /* === Решаем, нужен ли вообще РЕЕСТР АКТОВ в этом запуске === */
  const needActsGrid = needsActsGrid(inVals);
  logTime(`needsActsGrid = ${needActsGrid}`);

  /* === Справочник статей === */
  const dictIdx = buildDictionaryIndex_(shDict);
  logTime('buildDictionaryIndex_ завершена');
  const pairs  = dictIdx.pairs;
  const acts   = dictIdx.acts;
  const hashes = dictIdx.hashes;
  const meta   = dictIdx.meta;
  const byDec  = dictIdx.byDec;

  /* === Дубли по последним 50 строкам ПРОВОДОК === */
  const existing = buildExistingEntriesSet(shProv, 50, tz);
  logTime('buildExistingEntriesSet завершена');

  /* === РЕЕСТР АКТОВ (только ключи и флаги, без сумм и остатков) === */
  let actsGrid = null;
  let keyToRow = {};

  if (needActsGrid && shActs && shActs.getLastRow() > 1) {
    const actsIdx = buildActsIndex_(shActs);
    actsGrid = actsIdx.actsGrid;
    keyToRow = actsIdx.keyToRow;
    logTime('buildActsIndex_ завершена');
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

  // 🚀 ОПТИМИЗАЦИЯ: предварительный сбор всех вопросов для пользователя
  const questionsCache = {}; // key → boolean (ответ пользователя)
  
  // 🚀 КРИТИЧЕСКАЯ ОПТИМИЗАЦИЯ: пропускаем предпроверку в auto режиме
  // или если нет существующих проводок (нечего проверять на дубли)
  if (!auto && existing.size > 0) {
    const duplicateQuestions = [];
    const actFlagQuestions = [];
    
    let validatedCount = 0;
    let skippedBlank = 0;
    let skippedNoAmount = 0;
    let skippedInvalid = 0;

    for (let i = 0; i < inVals.length; i++) {
      const r = inVals[i];
      
      // 🚀 Быстрая проверка пустой строки (одна операция вместо .every())
      if (!r[2]) { // Если нет суммы (колонка D) — строка пустая или невалидная
        skippedBlank++;
        continue;
      }

      const hasAmount = r[2] !== '' && r[2] != null && isFinite(Number(r[2])) && Number(r[2]) !== 0;
      if (!hasAmount) {
        skippedNoAmount++;
        continue;
      }

      const basic = validateRowBasic(r, i);
      if (!basic.ok) {
        skippedInvalid++;
        continue;
      }
      
      validatedCount++;

      let { date, wallet, amount, article, decoding, act } = basic;
      if (basic.wantsToday) date = new Date();

      const key = `${fmtDate(date, tz)}|${article}|${decoding}|${amount}`;
      const isMasterOrRetention = (article === '% Мастер' || article === 'Возврат удержания');

      // Проверка дублей (только для обычных проводок)
      if (!isMasterOrRetention && existing.has(key)) {
        duplicateQuestions.push({ i, date, article, decoding, amount, key });
      }

      // Проверка повторных операций по актам
      if (isMasterOrRetention && actsGrid) {
        const actKey = makeActKey(decoding, act);
        const info = findActRowByKey_(actsGrid, keyToRow, actKey);
        if (!info.error) {
          const isMaster = (article === '% Мастер');
          const alreadyFlag = isMaster ? info.master : info.ret;
          if (alreadyFlag) {
            actFlagQuestions.push({ i, actKey });
          }
        }
      }
    }

    // Задаём вопросы про дубли (если есть)
    if (duplicateQuestions.length > 0) {
      let msg = 'Обнаружены дубли:\n\n';
      duplicateQuestions.forEach((q, idx) => {
        msg += `${idx + 1}. ${fmtDate(q.date, tz)} | ${q.article} | ${q.decoding} | ${q.amount}\n`;
      });
      msg += '\nВнести все повторно?';
      
      const answerAll = confirmDialog_('Дубликаты проводок', msg);
      duplicateQuestions.forEach(q => {
        questionsCache[`duplicate_${q.key}`] = answerAll;
      });
    }

    // Задаём вопросы про повторные операции по актам (если есть)
    if (actFlagQuestions.length > 0) {
      let msg = 'Обнаружены повторные выплаты по актам:\n\n';
      actFlagQuestions.forEach((q, idx) => {
        msg += `${idx + 1}. Строка B${10 + q.i}: акт ${q.actKey}\n`;
      });
      msg += '\nПовторить все операции?';
      
      const answerAll = confirmDialog_('Повторные операции по актам', msg);
      actFlagQuestions.forEach(q => {
        questionsCache[`act_flag_${q.actKey}`] = answerAll;
      });
    }
    
    console.log(`Предпроверка: обработано ${validatedCount} строк (пропущено: ${skippedBlank} пустых, ${skippedNoAmount} без суммы, ${skippedInvalid} невалидных)`);
  }
  logTime('предварительная проверка вопросов завершена');

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
      // 🚀 ОПТИМИЗАЦИЯ: используем кэш ответов
      const cacheKey = `duplicate_${key}`;
      const cached = questionsCache[cacheKey];
      if (cached === false) {
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
    const isMaster = (article === '% Мастер');
    const isRetention = (article === 'Возврат удержания');

    if (isMaster || isRetention) {
      const actResult = processActsRelatedEntry({
        article, decoding, act, shActs, actsGrid, keyToRow
      });

      if (!actResult.ok) {
        err(i, actResult.error);
        return;
      }

      const { alreadyFlag, targetCol, gridIndex, targetRow } = actResult;

      if (alreadyFlag) {
        if (auto) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
        // 🚀 ОПТИМИЗАЦИЯ: используем кэш ответов
        const actKey = makeActKey(decoding, act);
        const cacheKey = `act_flag_${actKey}`;
        const cached = questionsCache[cacheKey];
        if (cached === false) {
          err(i, 'Отменено: по этому акту уже стояла галочка выплаты');
          return;
        }
      }

      actsGrid[gridIndex][targetCol - 1] = true;
      if (isMaster) masterFlagRows.add(targetRow);
      else depFlagRows.add(targetRow);
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
  logTime(`обработка ${toWrite.length} строк завершена`);

  /* === Запись в ☑️ ПРОВОДКИ === */
  if (toWrite.length) {
    const curFilter = shProv.getFilter();
    if (curFilter) curFilter.remove();

    const start = findStartRowForProv_(shProv);
    const writeResult = writeProvodkiToSheet(shProv, toWrite, start);

    if (!writeResult.ok) {
      okDialog_('Ошибка', writeResult.error);
      return;
    }

    // Нативный быстрый Toast больше не показываем здесь — единый итоговый toast будет в конце
  }

  /* === Очистка/сохранение вводимых строк в ⏬ ВНЕСЕНИЕ ===
     — Чистим диапазон B10:G40 (содержимое и форматирование) для проведённых строк
     — Возвращаем только те строки, которые НЕ были проведены (B..G остаются)
  */
  clearProcessedInputRows(shIn, inVals, processedRows);

  /* === Новые расшифровки === */
  const newDecs = addNewDecodings_(shDict, toSuggest, meta, auto);

  /* === Запись флагов и стилей в РЕЕСТР АКТОВ === */
  if (shActs) {
    logTime('СТАРТ записи в РЕЕСТР АКТОВ');
    applyActsFlags_(shActs, masterFlagRows, depFlagRows);
    logTime('  └─ applyActsFlags_ завершена');
    applyRevenueColors_(shActs, revenueColorsByRow);
    logTime('  └─ applyRevenueColors_ завершена');
    applyStyleBlocks_(shActs, ACTS_COL.HANDS, masterFlagRows);
    logTime('  └─ applyStyleBlocks_ (HANDS) завершена');
    applyStyleBlocks_(shActs, ACTS_COL.DEPOSIT, depFlagRows);
    logTime('  └─ applyStyleBlocks_ (DEPOSIT) завершена');
    logTime('запись в РЕЕСТР АКТОВ завершена');
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
  
  logTime('ФИНИШ - все операции завершены'); // 🔍 ФИНАЛЬНЫЙ ЛОГ
  const totalTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  console.log(`🏁 Общее время выполнения: ${totalTime}s`);
  
  SpreadsheetApp.getActive().toast(summary + ` (${totalTime}s)`, 'Готово', 8);

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
// === Formatting and UI helpers moved to 12-formatting.js ===
// colorRows_, allowedWallets_, handleInternalTransfer_,
// pickArticleInteractive_, resolveArticleByDec_ → см. 12-formatting.js

/**
 * Ищет первую полностью пустую строку в блоке B10:G40 на листе "⏬ ВНЕСЕНИЕ".
 * Пустая = все ячейки B..G === '' / null / пробелы.
 * Возвращает номер строки или null, если нет.
 */
// findFirstEmptyRowInInput_ moved to 70-createFromActs.js

