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

    // Найти первую пустую строку по колонке B (Название), начиная со строки 2
    const nameCol = sheet.getRange(2, COL.NAME + 1, sheet.getMaxRows() - 1, 1).getValues();
    let targetRow = 2;
    for (let i = 0; i < nameCol.length; i++) {
      if (!nameCol[i][0]) { targetRow = i + 2; break; }
    }

    sheet.getRange(targetRow, 1, 1, HEADERS_SUBSCRIPTIONS.length).setValues([[
      newId,
      formData.name,
      formData.category,
      parseFloat(formData.amount),
      formData.currency,
      formData.period,
      nextDate,
      formData.status || 'Активна',
      '',                               // Последняя оплата
      false,                            // Оплачено
      formData.notifications !== false, // Уведомления
      parseInt(formData.remindDays) || 3,
      formData.payer || '',
      formData.payMethod || '',
      formData.notes || '',
      '',                               // Сумма/мес (formula)
      '',                               // Дней до оплаты (formula)
      ''                                // Calendar Event ID
    ]]);

    const lastRow = targetRow;
    sheet.getRange(lastRow, COL.MONTHLY_COST + 1).setFormula(
      '=IF(H' + lastRow + '="Активна";' +
      'IF(ISNUMBER(VALUE(LEFT(F' + lastRow + ';LEN(F' + lastRow + ')-5)));' +
      'D' + lastRow + '/VALUE(LEFT(F' + lastRow + ';LEN(F' + lastRow + ')-5));' +
      'SWITCH(F' + lastRow + ';"Месяц";D' + lastRow + ';"Квартал";D' + lastRow + '/3;' +
      '"Полгода";D' + lastRow + '/6;"Год";D' + lastRow + '/12;"Неделя";D' + lastRow + '*4,33;0));0)'
    );
    sheet.getRange(lastRow, COL.DAYS_UNTIL + 1).setFormula(
      '=IF(AND(H' + lastRow + '="Активна"; G' + lastRow + '<>""); G' + lastRow + '-TODAY(); "")'
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
