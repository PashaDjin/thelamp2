// 20-ui-dialogs.js — HTML-диалоги и UI-хелперы

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

function setDialogResult(token, data) {
  try {
    CacheService.getScriptCache().put(token, JSON.stringify(data || {}), 120);
  } catch (e) {
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
