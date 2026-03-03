// ==================== Setup.gs ====================
// Первоначальная настройка: создание листов, форматирование, триггеры

/**
 * Первоначальная настройка всей системы
 */
function initialSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Создать листы
    const subSheet = createSheetIfNotExists_(ss, SHEET_SUBSCRIPTIONS);
    const histSheet = createSheetIfNotExists_(ss, SHEET_HISTORY);
    const settSheet = createSheetIfNotExists_(ss, SHEET_SETTINGS);
    const lookSheet = createSheetIfNotExists_(ss, SHEET_LOOKUPS);

    // 2. Заполнить справочники
    setupLookups_(lookSheet);

    // 3. Заполнить настройки
    setupSettings_(settSheet);

    // 4. Настроить лист подписок
    setupSubscriptionsSheet_(ss, subSheet);

    // 5. Настроить лист истории
    setupHistorySheet_(histSheet);

    // 6. Создать календарь
    const settings = getSettings();
    const calendarName = settings['Название календаря'] || '💳 Подписки';
    getOrCreateCalendar(calendarName);

    // 7. Настроить триггеры
    setupTriggers_();

    ui.alert('✅ Настройка завершена! Календарь создан, триггеры установлены.');
  } catch (e) {
    console.error('initialSetup error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка настройки: ' + e.message);
  }
}

/**
 * Создать лист, если он не существует
 */
function createSheetIfNotExists_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * Заполнить лист "Справочники"
 */
function setupLookups_(sheet) {
  sheet.clear();

  // Заголовки
  sheet.getRange(1, 1, 1, HEADERS_LOOKUPS.length).setValues([HEADERS_LOOKUPS]);
  sheet.getRange(1, 1, 1, HEADERS_LOOKUPS.length)
    .setFontWeight('bold')
    .setBackground('#4A86C8')
    .setFontColor('#FFFFFF');

  // Данные
  const maxRows = Math.max(
    LOOKUP_CATEGORIES.length,
    LOOKUP_PERIODS.length,
    LOOKUP_CURRENCIES.length,
    LOOKUP_PAY_METHODS.length,
    LOOKUP_STATUSES.length
  );

  for (let i = 0; i < maxRows; i++) {
    const row = i + 2;
    if (i < LOOKUP_CATEGORIES.length) sheet.getRange(row, 1).setValue(LOOKUP_CATEGORIES[i]);
    if (i < LOOKUP_PERIODS.length) sheet.getRange(row, 2).setValue(LOOKUP_PERIODS[i]);
    if (i < LOOKUP_CURRENCIES.length) sheet.getRange(row, 3).setValue(LOOKUP_CURRENCIES[i]);
    if (i < LOOKUP_PAY_METHODS.length) sheet.getRange(row, 4).setValue(LOOKUP_PAY_METHODS[i]);
    if (i < LOOKUP_STATUSES.length) sheet.getRange(row, 5).setValue(LOOKUP_STATUSES[i]);
  }
}

/**
 * Заполнить лист "Настройки"
 */
function setupSettings_(sheet) {
  // Не перезаписываем, если уже есть данные
  const existing = sheet.getDataRange().getValues();
  if (existing.length > 1 && existing[0][0]) return;

  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Параметр', 'Значение']]);
  sheet.getRange(1, 1, 1, 2)
    .setFontWeight('bold')
    .setBackground('#4A86C8')
    .setFontColor('#FFFFFF');

  for (let i = 0; i < SETTINGS_KEYS.length; i++) {
    sheet.getRange(i + 2, 1).setValue(SETTINGS_KEYS[i]);
    sheet.getRange(i + 2, 2).setValue(SETTINGS_DEFAULTS[i]);
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 300);
}

/**
 * Настроить лист "Подписки": заголовки, форматирование, валидация, формулы
 */
function setupSubscriptionsSheet_(ss, sheet) {
  // Заголовки (только если лист пустой)
  const existing = sheet.getDataRange().getValues();
  if (existing.length <= 1 || !existing[0][1]) {
    sheet.getRange(1, 1, 1, HEADERS_SUBSCRIPTIONS.length).setValues([HEADERS_SUBSCRIPTIONS]);
  }

  // Стиль заголовков
  const headerRange = sheet.getRange(1, 1, 1, HEADERS_SUBSCRIPTIONS.length);
  headerRange.setFontWeight('bold')
    .setBackground('#4A86C8')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // Закрепить заголовок
  sheet.setFrozenRows(1);

  // Ширины колонок
  for (const [col, width] of Object.entries(COLUMN_WIDTHS)) {
    sheet.setColumnWidth(parseInt(col) + 1, width);
  }

  // Скрыть колонку R (Calendar Event ID)
  sheet.hideColumns(COL.CALENDAR_ID + 1);

  // Настроить чекбоксы (J — Оплачено, K — Уведомления) для строк 2-100
  const checkboxRange1 = sheet.getRange(2, COL.IS_PAID + 1, 99, 1);
  const checkboxRange2 = sheet.getRange(2, COL.NOTIFICATIONS + 1, 99, 1);
  checkboxRange1.insertCheckboxes();
  checkboxRange2.insertCheckboxes();

  // Валидация — используем диапазоны из Справочников
  const lookSheet = ss.getSheetByName(SHEET_LOOKUPS);
  setupDropdownValidation_(sheet, lookSheet);

  // Формулы для строк 2-100
  for (let row = 2; row <= 100; row++) {
    // P: Сумма/мес
    sheet.getRange(row, COL.MONTHLY_COST + 1).setFormula(
      '=IF(H' + row + '="Активна";' +
      'IF(ISNUMBER(VALUE(LEFT(F' + row + ';LEN(F' + row + ')-5)));' +
      'D' + row + '/VALUE(LEFT(F' + row + ';LEN(F' + row + ')-5));' +
      'SWITCH(F' + row + ';"Месяц";D' + row + ';"Квартал";D' + row + '/3;' +
      '"Полгода";D' + row + '/6;"Год";D' + row + '/12;"Неделя";D' + row + '*4,33;0));0)'
    );
    // Q: Дней до оплаты
    sheet.getRange(row, COL.DAYS_UNTIL + 1).setFormula(
      '=IF(AND(H' + row + '="Активна"; G' + row + '<>""); G' + row + '-TODAY(); "")'
    );
  }

  // Формат даты для колонок G и I
  sheet.getRange(2, COL.NEXT_DATE + 1, 99, 1).setNumberFormat('dd.MM.yyyy');
  sheet.getRange(2, COL.LAST_PAID + 1, 99, 1).setNumberFormat('dd.MM.yyyy');

  // Условное форматирование
  setupConditionalFormatting_(sheet);
}

/**
 * Настроить выпадающие списки с валидацией
 */
function setupDropdownValidation_(subSheet, lookSheet) {
  const lastRow = 100;

  // C — Категория (из Справочники!A2:A)
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lookSheet.getRange('A2:A' + (LOOKUP_CATEGORIES.length + 1)), true)
    .setAllowInvalid(false)
    .build();
  subSheet.getRange(2, COL.CATEGORY + 1, lastRow - 1, 1).setDataValidation(catRule);

  // E — Валюта
  const curRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lookSheet.getRange('C2:C' + (LOOKUP_CURRENCIES.length + 1)), true)
    .setAllowInvalid(false)
    .build();
  subSheet.getRange(2, COL.CURRENCY + 1, lastRow - 1, 1).setDataValidation(curRule);

  // F — Период
  const perRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lookSheet.getRange('B2:B' + (LOOKUP_PERIODS.length + 1)), true)
    .setAllowInvalid(true)  // Разрешить произвольные периоды вида "N мес."
    .build();
  subSheet.getRange(2, COL.PERIOD + 1, lastRow - 1, 1).setDataValidation(perRule);

  // H — Статус
  const statRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lookSheet.getRange('E2:E' + (LOOKUP_STATUSES.length + 1)), true)
    .setAllowInvalid(false)
    .build();
  subSheet.getRange(2, COL.STATUS + 1, lastRow - 1, 1).setDataValidation(statRule);

  // N — Способ оплаты
  const payRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(lookSheet.getRange('D2:D' + (LOOKUP_PAY_METHODS.length + 1)), true)
    .setAllowInvalid(false)
    .build();
  subSheet.getRange(2, COL.PAY_METHOD + 1, lastRow - 1, 1).setDataValidation(payRule);
}

/**
 * Настроить условное форматирование на листе "Подписки"
 */
function setupConditionalFormatting_(sheet) {
  // Удалить существующие правила
  sheet.clearConditionalFormatRules();

  const range = sheet.getRange('A2:R100');
  const rules = [];

  // 1. Отменена — зачёркнуто, серый текст
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="Отменена"')
    .setStrikethrough(true)
    .setFontColor('#999999')
    .setRanges([range])
    .build());

  // 2. Приостановлена — курсив, светло-серый
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2="Приостановлена"')
    .setItalic(true)
    .setFontColor('#AAAAAA')
    .setRanges([range])
    .build());

  // 3. Оплачено = TRUE — зелёный фон
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$J2=TRUE')
    .setBackground('#D9EAD3')
    .setRanges([range])
    .build());

  // 4. Просрочено (дней <= 0) — красный фон, жирный
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2="Активна",$Q2<=0)')
    .setBackground('#F4C7C3')
    .setBold(true)
    .setRanges([range])
    .build());

  // 5. Скоро (дней <= 3) — оранжевый фон
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2="Активна",$Q2<=3,$Q2>0)')
    .setBackground('#FCE8B2')
    .setRanges([range])
    .build());

  // 6. На подходе (дней <= 7) — жёлтый фон
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H2="Активна",$Q2<=7,$Q2>3)')
    .setBackground('#FFF2CC')
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * Настроить лист "История оплат"
 */
function setupHistorySheet_(sheet) {
  const existing = sheet.getDataRange().getValues();
  if (existing.length <= 1 || !existing[0][0]) {
    sheet.getRange(1, 1, 1, HEADERS_HISTORY.length).setValues([HEADERS_HISTORY]);
  }

  sheet.getRange(1, 1, 1, HEADERS_HISTORY.length)
    .setFontWeight('bold')
    .setBackground('#4A86C8')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(1);
  sheet.getRange(2, 4, 999, 1).setNumberFormat('dd.MM.yyyy');
}

/**
 * Настроить триггеры (удалить старые, создать новые)
 */
function setupTriggers_() {
  // Удалить существующие триггеры этого проекта
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'dailyCheck' || handler === 'onEditTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // Ежедневный триггер для dailyCheck
  ScriptApp.newTrigger('dailyCheck')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // Installable onEdit триггер
  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}
