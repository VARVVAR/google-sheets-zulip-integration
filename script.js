/**
 * Автоматизация уведомлений в Zulip при изменении статуса в таблице.
 */

const ZULIP_CONFIG = {
  URL: "YOUR_URL_ZULIP",
  EMAIL: "YOUR_EMAIL",
  API_KEY: "YOUR_ZULIP_API_KEY_HERE"
};

const COLUMNS = {
  OBJECT: 1,      // A
  TASK: 2,        // B
  REQUEST_ID: 3,  // C
  DATE_READY: 5,  // E
  DATE_SENT: 6,   // F
  TRIGGER: 9,     // I (Чекбокс)
  RESPONSIBLE: 10 // J (Фамилия)
};

function sendToZulip(e) {
  if (!e || !e.range) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // Реагируем только на нажатие чекбокса в нужной колонке
  if (col !== COLUMNS.TRIGGER || range.getValue() !== true) return;

  // Карта пользователей: Фамилия в таблице -> ID в Zulip
  const userMap = {
    "NAME": ID zulip,  
  };

  const responsibleName = sheet.getRange(row, COLUMNS.RESPONSIBLE).getValue();
  const userId = userMap[responsibleName];

  if (!userId) {
    Browser.msgBox("⚠️ Ошибка: Не указан ответственный в колонке J или фамилия не найдена в базе.");
    range.setValue(false);
    return;
  }

  // Сбор данных из строки
  const rowData = sheet.getRange(row, 1, 1, 10).getValues()[0];
  const objectName = rowData[COLUMNS.OBJECT - 1];
  const taskDesc   = rowData[COLUMNS.TASK - 1];
  const reqNumber  = rowData[COLUMNS.REQUEST_ID - 1];

  // Форматирование даты для читабельности в мессенджере
  const formatDate = (date) => {
    return (date instanceof Date) ? Utilities.formatDate(date, "GMT+3", "dd.MM") : (date || "—");
  };

  const datePrep = formatDate(rowData[COLUMNS.DATE_READY - 1]);
  const dateSent = formatDate(rowData[COLUMNS.DATE_SENT - 1]);

  // Формируем красивый Markdown-текст
  const messageContent = [
    `✅ **Техника готова**`,
    `📍 **Объект:** ${objectName}`,
    `📌 **Что:** ${taskDesc}`,
    `📄 **Заявка №:** ${reqNumber}`,
    `🛠 **Подготовлено:** ${datePrep}`,
    `🚚 **Уехало:** ${dateSent}`
  ].join('\n');

  try {
    const payload = {
      "type": "private",
      "to": JSON.stringify([userId]), 
      "content": messageContent
    };

    const authHeader = "Basic " + Utilities.base64Encode(`${ZULIP_CONFIG.EMAIL}:${ZULIP_CONFIG.API_KEY}`);

    const response = UrlFetchApp.fetch(ZULIP_CONFIG.URL, {
      "method": "post",
      "payload": payload,
      "headers": { "Authorization": authHeader },
      "muteHttpExceptions": true
    });

    const result = JSON.parse(response.getContentText());

    if (result.result === "success") {
      // Визуальное подтверждение отправки в таблице
      range.clearDataValidations() // Убираем чекбокс
           .setValue("✅ ОТПРАВЛЕНО")
           .setFontColor("#274e13")
           .setFontWeight("bold")
           .setHorizontalAlignment("center");

      sheet.getParent().toast(`Уведомление для ${responsibleName} успешно доставлено.`);
    } else {
      throw new Error(result.msg || "Неизвестная ошибка API");
    }

  } catch (err) {
    console.error("Zulip Integration Error: " + err.message);
    Browser.msgBox("❌ Ошибка при отправке: " + err.message);
    range.setValue(false);
  }
}
