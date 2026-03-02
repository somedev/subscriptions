// ==================== DailyCheck.gs ====================
// Ежедневная проверка: напоминания, safety net для чекбоксов

/**
 * Ежедневная проверка подписок (вызывается триггером)
 */
function dailyCheck() {
  try {
    const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const props = PropertiesService.getScriptProperties();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[COL.NAME]) continue;

      const status = row[COL.STATUS];
      if (status !== 'Активна') continue;

      const rowIndex = i + 1;

      // Safety net: если чекбокс "Оплачено" = TRUE, а onEdit не сработал
      if (row[COL.IS_PAID] === true) {
        try {
          confirmPayment(rowIndex);
        } catch (e) {
          console.error('dailyCheck confirmPayment error for row ' + rowIndex + ': ' + e.message);
        }
        continue;
      }

      // Проверка напоминаний
      if (row[COL.NOTIFICATIONS] !== true) continue;

      const nextDate = row[COL.NEXT_DATE];
      if (!nextDate) continue;

      const payDate = new Date(nextDate);
      payDate.setHours(0, 0, 0, 0);

      const diffMs = payDate.getTime() - today.getTime();
      const daysUntil = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
      const remindDays = row[COL.REMIND_DAYS] || 3;

      if (daysUntil <= remindDays && daysUntil >= 0) {
        // Проверить, не отправляли ли уже напоминание для этого цикла
        const reminderKey = 'reminder_' + row[COL.ID] + '_' + formatDate(nextDate);
        const lastReminder = props.getProperty(reminderKey);

        if (!lastReminder) {
          sendReminder(row, daysUntil);
          props.setProperty(reminderKey, new Date().toISOString());
        }
      }
    }
  } catch (e) {
    console.error('dailyCheck error: ' + e.message);
  }
}
