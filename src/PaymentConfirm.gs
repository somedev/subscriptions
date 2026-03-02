// ==================== PaymentConfirm.gs ====================
// Подтверждение оплаты, логирование в историю, расчёт следующей даты

/**
 * Подтвердить оплату подписки по номеру строки
 */
function confirmPayment(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SUBSCRIPTIONS);
    const row = sheet.getRange(rowIndex, 1, 1, HEADERS_SUBSCRIPTIONS.length).getValues()[0];

    const name = row[COL.NAME];
    const amount = row[COL.AMOUNT];
    const nextDate = row[COL.NEXT_DATE];

    // Валидация
    if (!name) {
      throw new Error('Строка ' + rowIndex + ' не содержит названия подписки.');
    }
    if (!amount) {
      throw new Error('Для "' + name + '" не указана сумма.');
    }
    if (!nextDate) {
      throw new Error('Для "' + name + '" не указана дата следующей оплаты.');
    }

    const subId = row[COL.ID];
    const currency = row[COL.CURRENCY];
    const period = row[COL.PERIOD];

    // 1. Записать в историю оплат
    logPayment_(subId, name, amount, currency);

    // 2. Рассчитать следующую дату оплаты
    const newDate = calculateNextDate(nextDate, period);

    // 3. Обновить строку подписки
    sheet.getRange(rowIndex, COL.NEXT_DATE + 1).setValue(newDate);
    sheet.getRange(rowIndex, COL.LAST_PAID + 1).setValue(new Date());
    sheet.getRange(rowIndex, COL.IS_PAID + 1).setValue(false);

    // 4. Обновить событие в календаре
    try {
      updateCalendarEventForRow_(sheet, rowIndex, row, newDate);
    } catch (calError) {
      console.error('Ошибка обновления календаря: ' + calError.message);
    }

    // 5. Показать уведомление
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ Оплата "' + name + '" подтверждена. Следующая: ' + formatDate(newDate),
      'Оплата',
      5
    );
  } catch (e) {
    console.error('confirmPayment error: ' + e.message);
    throw e;
  }
}

/**
 * Записать оплату в лист "История оплат"
 */
function logPayment_(subId, name, amount, currency) {
  const histSheet = getSheetByName(SHEET_HISTORY);
  const newId = getNextId(histSheet);
  const payer = Session.getActiveUser().getEmail() || '';

  histSheet.appendRow([
    newId,
    subId,
    name,
    new Date(),
    amount,
    currency,
    payer,
    ''
  ]);
}

/**
 * Обновить событие в календаре для строки после оплаты
 */
function updateCalendarEventForRow_(sheet, rowIndex, row, newDate) {
  const settings = getSettings();
  const calendarName = settings['Название календаря'] || '💳 Подписки';
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length === 0) return;

  const calendar = calendars[0];
  const eventId = row[COL.CALENDAR_ID];

  // Удалить старое событие
  if (eventId) {
    try {
      calendar.getEventById(eventId).deleteEvent();
    } catch (e) {
      // Событие уже удалено — игнорируем
    }
  }

  // Создать новое событие
  const name = row[COL.NAME];
  const amount = row[COL.AMOUNT];
  const currency = row[COL.CURRENCY];
  const remindDays = row[COL.REMIND_DAYS] || 3;

  const title = '💳 ' + name + ' — ' + amount + ' ' + currency;
  const description = buildEventDescription_(row);

  const event = calendar.createAllDayEvent(title, newDate);
  event.setDescription(description);
  event.removeAllReminders();
  event.addPopupReminder(remindDays * 24 * 60); // За N дней
  event.addPopupReminder(0); // В день оплаты

  // Сохранить новый Event ID
  sheet.getRange(rowIndex, COL.CALENDAR_ID + 1).setValue(event.getId());
}

/**
 * Сформировать описание события для Google Calendar
 */
function buildEventDescription_(row) {
  return '📌 Подписка: ' + row[COL.NAME] +
    '\n💰 Сумма: ' + row[COL.AMOUNT] + ' ' + row[COL.CURRENCY] +
    '\n📅 Период: ' + row[COL.PERIOD] +
    '\n👤 Кто платит: ' + (row[COL.PAYER] || '—') +
    '\n💳 Способ: ' + (row[COL.PAY_METHOD] || '—') +
    '\n📝 ' + (row[COL.NOTES] || '') +
    '\n\n📊 Таблица: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

/**
 * Installable onEdit триггер
 */
function onEditTrigger(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_SUBSCRIPTIONS) return;

    const col = e.range.getColumn();
    const row = e.range.getRow();

    // Колонка J (Оплачено) — индекс 10 (1-based)
    if (col !== COL.IS_PAID + 1) return;
    if (row < 2) return;
    if (e.value !== 'TRUE') return;

    confirmPayment(row);
  } catch (e) {
    console.error('onEditTrigger error: ' + e.message);
  }
}
