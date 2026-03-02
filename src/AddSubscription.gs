// ==================== AddSubscription.gs ====================
// Сайдбар для добавления новой подписки

/**
 * Показать боковую панель для добавления подписки
 */
function showAddDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddSubscriptionForm')
    .setTitle('➕ Новая подписка')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Получить данные для выпадающих списков в форме
 */
function getFormData() {
  const settings = getSettings();
  const familyStr = settings['Члены семьи'] || '';
  const family = familyStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  const defaultCurrency = settings['Валюта по умолчанию'] || 'RUB';

  return {
    categories: LOOKUP_CATEGORIES,
    periods: LOOKUP_PERIODS,
    currencies: LOOKUP_CURRENCIES,
    payMethods: LOOKUP_PAY_METHODS,
    statuses: LOOKUP_STATUSES,
    family: family,
    defaultCurrency: defaultCurrency
  };
}

/**
 * Обработать данные формы добавления подписки
 */
function processAddForm(formData) {
  try {
    const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
    const newId = getNextId(sheet);

    const nextDate = new Date(formData.nextDate);

    sheet.appendRow([
      newId,
      formData.name,
      formData.category,
      parseFloat(formData.amount),
      formData.currency,
      formData.period,
      nextDate,
      formData.status || 'Активна',
      '',                              // Последняя оплата
      false,                           // Оплачено
      formData.notifications !== false, // Уведомления
      parseInt(formData.remindDays) || 3,
      formData.payer || '',
      formData.payMethod || '',
      formData.notes || '',
      '',                              // Сумма/мес (formula)
      '',                              // Дней до оплаты (formula)
      ''                               // Calendar Event ID
    ]);

    // Установить формулы для новой строки
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, COL.MONTHLY_COST + 1).setFormula(
      '=IF(H' + lastRow + '="Активна", SWITCH(F' + lastRow + ', "Месяц",D' + lastRow +
      ', "Квартал",D' + lastRow + '/3, "Полгода",D' + lastRow + '/6, "Год",D' + lastRow +
      '/12, "Неделя",D' + lastRow + '*4.33, 0), 0)'
    );
    sheet.getRange(lastRow, COL.DAYS_UNTIL + 1).setFormula(
      '=IF(AND(H' + lastRow + '="Активна", G' + lastRow + '<>""), G' + lastRow + '-TODAY(), "")'
    );

    // Чекбоксы
    sheet.getRange(lastRow, COL.IS_PAID + 1).insertCheckboxes();
    sheet.getRange(lastRow, COL.NOTIFICATIONS + 1).insertCheckboxes();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ Подписка "' + formData.name + '" добавлена', 'Добавлено', 3
    );

    return { success: true };
  } catch (e) {
    console.error('processAddForm error: ' + e.message);
    return { success: false, error: e.message };
  }
}
