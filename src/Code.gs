// ==================== Code.gs ====================
// Точка входа: onOpen(), doGet(), пользовательское меню

/**
 * Web App entry point — мобильный интерфейс
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('💳 Подписки')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Включение HTML-файлов (CSS, JS) в шаблон
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Создание пользовательского меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💳 Подписки')
    .addItem('✅ Отметить оплату', 'menuConfirmPayment')
    .addItem('📅 Синхронизировать календарь', 'syncCalendar')
    .addItem('📊 Обновить статистику', 'updateStatistics')
    .addItem('➕ Добавить подписку', 'showAddDialog')
    .addSeparator()
    .addItem('⚙️ Первоначальная настройка', 'initialSetup')
    .addToUi();
}

/**
 * Обработка пункта меню "Отметить оплату" — для текущей выделенной строки
 */
function menuConfirmPayment() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== SHEET_SUBSCRIPTIONS) {
      SpreadsheetApp.getUi().alert('⚠️ Перейдите на лист "Подписки" и выберите строку.');
      return;
    }
    const row = SpreadsheetApp.getActiveRange().getRow();
    if (row < 2) {
      SpreadsheetApp.getUi().alert('⚠️ Выберите строку с подпиской (не заголовок).');
      return;
    }
    confirmPayment(row);
  } catch (e) {
    console.error('menuConfirmPayment error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + e.message);
  }
}
