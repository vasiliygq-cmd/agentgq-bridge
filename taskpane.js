/* AgentGQ Bridge Test — client-side only; no mail content is transmitted. */
(function () {
  'use strict';
  const status = document.getElementById('status');
  const button = document.getElementById('read');
  const result = document.getElementById('result');
  const metadata = document.getElementById('metadata');

  function showStatus(message, isError) {
    status.textContent = message;
    status.className = isError ? 'error' : 'ok';
  }

  function readMessage() {
    const mailbox = window.Office && Office.context && Office.context.mailbox;
    const item = mailbox && mailbox.item;
    if (!item || !item.body || typeof item.body.getAsync !== 'function') {
      showStatus('Не удалось получить доступ к открытому письму. Проверьте поддержку Office.js в этом Outlook.', true);
      return;
    }
    button.disabled = true;
    showStatus('Читаем открытое письмо…', false);
    try {
      item.body.getAsync(Office.CoercionType.Text, function (response) {
        button.disabled = false;
        if (response.status !== Office.AsyncResultStatus.Succeeded) {
          const msg = response.error && response.error.message ? response.error.message : 'Неизвестная ошибка Office.js';
          showStatus('Ошибка чтения: ' + msg, true);
          return;
        }
        result.textContent = response.value || '(Тело письма пустое)';
        const subject = item.subject || '(без темы)';
        metadata.textContent = 'Тема: ' + subject + ' · Символов: ' + (response.value || '').length;
        showStatus('Письмо прочитано локально. Данные никуда не отправлялись.', false);
      });
    } catch (err) {
      button.disabled = false;
      showStatus('Ошибка: ' + (err && err.message ? err.message : String(err)), true);
    }
  }

  button.addEventListener('click', readMessage);
  if (!window.Office || typeof Office.onReady !== 'function') {
    showStatus('Библиотека Office.js не загрузилась. Проверьте доступ к appsforoffice.microsoft.com.', true);
    return;
  }
  Office.onReady(function (info) {
    if (!info || info.host !== Office.HostType.Outlook) {
      showStatus('Панель открыта вне Outlook. Запустите её из письма в OWA.', true);
      return;
    }
    showStatus('Office.js готов. Откройте письмо и нажмите кнопку.', false);
    button.disabled = false;
  }).catch(function (err) {
    showStatus('Ошибка инициализации Office.js: ' + (err && err.message ? err.message : String(err)), true);
  });
})();
