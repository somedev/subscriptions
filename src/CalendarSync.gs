// ==================== CalendarSync.gs ====================
// Синхронизация подписок с Google Calendar

/**
 * Полная синхронизация календаря со всеми подписками
 */
function syncCalendar() {
  try {
    const settings = getSettings();
    const calendarName = settings['Название календаря'] || '💳 Подписки';
    const calendar = getOrCreateCalendar(calendarName);

    const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
    const data = sheet.getDataRange().getValues();

    let created = 0;
    let updated = 0;
    let deleted = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COL.NAME]) continue;

      const rowIndex = i + 1;
      const status = row[COL.STATUS];
      const eventId = row[COL.CALENDAR_ID];

      if (status === 'Активна') {
        const name = row[COL.NAME];
        const amount = row[COL.AMOUNT];
        const currency = row[COL.CURRENCY];
        const nextDate = row[COL.NEXT_DATE];
        const remindDays = row[COL.REMIND_DAYS] || 3;

        if (!nextDate) continue;

        const title = '💳 ' + name + ' — ' + amount + ' ' + currency;
        const description = buildEventDescription_(row);

        if (eventId) {
          // Обновить существующее событие
          try {
            const event = calendar.getEventById(eventId);
            if (event) {
              event.setTitle(title);
              event.setAllDayDate(nextDate);
              event.setDescription(description);
              event.removeAllReminders();
              event.addPopupReminder(remindDays * 24 * 60);
              event.addPopupReminder(0);
              updated++;
            } else {
              // Событие не найдено — создаём новое
              const newEvent = createCalendarEvent_(calendar, title, nextDate, description, remindDays);
              sheet.getRange(rowIndex, COL.CALENDAR_ID + 1).setValue(newEvent.getId());
              created++;
            }
          } catch (e) {
            // Ошибка — создаём заново
            console.error('Ошибка обновления события: ' + e.message);
            const newEvent = createCalendarEvent_(calendar, title, nextDate, description, remindDays);
            sheet.getRange(rowIndex, COL.CALENDAR_ID + 1).setValue(newEvent.getId());
            created++;
          }
        } else {
          // Создать новое событие
          const newEvent = createCalendarEvent_(calendar, title, nextDate, description, remindDays);
          sheet.getRange(rowIndex, COL.CALENDAR_ID + 1).setValue(newEvent.getId());
          created++;
        }
      } else {
        // Приостановлена или Отменена — удалить событие
        if (eventId) {
          try {
            const event = calendar.getEventById(eventId);
            if (event) {
              event.deleteEvent();
              deleted++;
            }
          } catch (e) {
            console.error('Ошибка удаления события: ' + e.message);
          }
          sheet.getRange(rowIndex, COL.CALENDAR_ID + 1).setValue('');
        }
      }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '📅 Синхронизация завершена. Создано: ' + created + ', обновлено: ' + updated + ', удалено: ' + deleted,
      'Календарь',
      5
    );
  } catch (e) {
    console.error('syncCalendar error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка синхронизации: ' + e.message);
  }
}

/**
 * Создать событие в календаре
 */
function createCalendarEvent_(calendar, title, date, description, remindDays) {
  const event = calendar.createAllDayEvent(title, date);
  event.setDescription(description);
  event.removeAllReminders();
  event.addPopupReminder(remindDays * 24 * 60);
  event.addPopupReminder(0);
  return event;
}
