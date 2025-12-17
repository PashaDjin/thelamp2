/**
 * ═══════════════════════════════════════════════════════════════════════════
 * 20-ui-dialogs.js — Диалоговые окна для взаимодействия с пользователем
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * Этот файл содержит функции для показа диалогов (всплывающих окон).
 * 
 * Почему HTML-диалоги, а не стандартные alert()?
 * - Более красивый внешний вид
 * - Можно добавить поля ввода
 * - Работают в Google Sheets без ограничений
 * 
 * Как это работает?
 * 1. Создаём HTML-страницу с кнопками
 * 2. Показываем её как модальное окно
 * 3. Ждём, пока пользователь нажмёт кнопку
 * 4. Сохраняем результат через CacheService
 * 5. Закрываем диалог и возвращаем результат
 * ═══════════════════════════════════════════════════════════════════════════
 */

/**
 * Показывает HTML-диалог и ждёт ответа пользователя
 * 
 * Это главная функция для всех диалогов. Она:
 * - Создаёт HTML с кнопками и опциональным полем ввода
 * - Показывает модальное окно (блокирует работу с таблицей)
 * - Ждёт, пока пользователь нажмёт кнопку (до 20 секунд)
 * - Возвращает результат: какую кнопку нажали и что ввели
 * 
 * @param {Object} options - Настройки диалога
 * @param {string} options.title - Заголовок окна
 * @param {string} options.message - Текст сообщения (можно многострочный)
 * @param {string[]} options.buttons - Массив названий кнопок, например ['Да', 'Нет']
 * @param {boolean} [options.withInput=false] - Добавить ли поле ввода текста
 * @param {string} [options.defaultValue=''] - Начальное значение в поле ввода
 * @returns {{button: string, value: string}|null} - Объект с нажатой кнопкой и введённым текстом, или null при таймауте
 */
function showDialogAndWait_({ title, message, buttons, withInput = false, defaultValue = '' }) {
  // Создаём уникальный токен для этого диалога (чтобы не перепутать результаты разных диалогов)
  const cache = CacheService.getScriptCache();
  const token = `dlg_${Date.now()}_${Math.random().toString(16).slice(2)}`;
  cache.remove(token); // Очищаем старые данные с таким же токеном (если были)

  // Создаём HTML-страницу с диалогом
  const html = HtmlService.createHtmlOutput(`
    <div style="white-space:pre-wrap;">${escapeHtml_(message)}</div>
    ${withInput ? `<div><input id="dlg-input" value="${escapeHtml_(defaultValue)}" /></div>` : ''}
    <div>${buttons.map(b => `<button onclick="submitDialog('${b}')">${escapeHtml_(b)}</button>`).join('')}</div>
    <script>
      // Функция вызывается при нажатии на кнопку
      function submitDialog(btn){
        // Берём значение из поля ввода (если оно есть)
        const v = document.getElementById('dlg-input') ? document.getElementById('dlg-input').value : '';
        // Сохраняем результат в кеш и закрываем окно
        google.script.run.withSuccessHandler(function(){ google.script.host.close(); })
          .setDialogResult('${token}', { button: btn, value: v });
      }
      // Автофокус на первую кнопку после загрузки
      document.addEventListener('DOMContentLoaded', function(){
        const b = document.querySelector('button'); if(b) b.focus();
      });
    </script>
  `)
    .setWidth(380)
    .setHeight(withInput ? 180 : 140);

  // Показываем диалог пользователю
  SpreadsheetApp.getUi().showModalDialog(html, title);

  // Ждём результат (макс 20 секунд)
  const timeoutMs = 20000;
  const start = Date.now();
  
  while (Date.now() - start < timeoutMs) {
    // Проверяем, сохранил ли пользователь результат
    const data = cache.get(token);
    if (data) {
      cache.remove(token); // Очищаем кеш
      try {
        return JSON.parse(data); // Возвращаем результат
      } catch (e) {
        return null;
      }
    }
    Utilities.sleep(30); // Ждём 30 мс перед следующей проверкой
  }

  // Если вышел таймаут — очищаем и возвращаем null
  cache.remove(token);
  return null;
}

/**
 * Сохраняет результат диалога в кеш
 * 
 * Вызывается из HTML-диалога при нажатии на кнопку.
 * Сохраняет данные так, чтобы функция showDialogAndWait_ могла их прочитать.
 * 
 * @param {string} token - Уникальный идентификатор диалога
 * @param {Object} data - Данные для сохранения (кнопка + введённый текст)
 */
function setDialogResult(token, data) {
  try {
    // Пытаемся сохранить в кеш (быстро, но ограничен 100 КБ)
    CacheService.getScriptCache().put(token, JSON.stringify(data || {}), 120);
  } catch (e) {
    // Если кеш переполнен — используем Properties (медленнее, но надёжнее)
    PropertiesService.getDocumentProperties().setProperty(token, JSON.stringify(data || {}));
  }
}

/**
 * Показывает диалог с вопросом "Да/Нет"
 * 
 * Удобная обёртка над showDialogAndWait_ для простых подтверждений.
 * 
 * Пример использования:
 *   if (confirmDialog_('Удаление', 'Точно удалить?')) {
 *     // Пользователь нажал "Да"
 *   }
 * 
 * @param {string} title - Заголовок диалога
 * @param {string} message - Текст вопроса
 * @returns {boolean} - true если нажали "Да", false если "Нет" или закрыли окно
 */
function confirmDialog_(title, message) {
  const res = showDialogAndWait_({ title, message, buttons: ['Да', 'Нет'] });
  return !!(res && res.button === 'Да');
}

/**
 * Показывает информационное сообщение с кнопкой "Ок"
 * 
 * Используется для уведомлений, где не нужен выбор.
 * 
 * Пример:
 *   okDialog_('Готово', 'Проводки успешно перенесены!');
 * 
 * @param {string} title - Заголовок
 * @param {string} message - Текст сообщения
 */
function okDialog_(title, message) {
  showDialogAndWait_({ title, message, buttons: ['Ок'] });
}

/**
 * Показывает диалог с полем для ввода текста
 * 
 * Используется когда нужно запросить у пользователя какое-то значение.
 * 
 * Пример:
 *   const result = promptDialog_('Название', 'Введите название статьи:', 'Зарплата');
 *   if (result.button === 'Ok') {
 *     const name = result.text; // Введённое пользователем значение
 *   }
 * 
 * @param {string} title - Заголовок диалога
 * @param {string} message - Текст подсказки
 * @param {string} defaultValue - Значение по умолчанию в поле ввода
 * @returns {{button: string, text: string}} - Объект с кнопкой ('Ok' или 'Cancel') и введённым текстом
 */
function promptDialog_(title, message, defaultValue) {
  const res = showDialogAndWait_({ title, message, buttons: ['Ок', 'Отмена'], withInput: true, defaultValue });
  if (!res || res.button !== 'Ок') return { button: 'Cancel', text: '' };
  return { button: 'Ok', text: res.value };
}
