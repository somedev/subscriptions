// ==================== Notifications.gs ====================
// Email-уведомления о предстоящих оплатах

/**
 * Отправить напоминание об оплате
 */
function sendReminder(subscriptionRow, daysUntil) {
  try {
    const settings = getSettings();
    const emailsStr = settings['Email для уведомлений'];
    if (!emailsStr) return;

    const emails = emailsStr.split(',').map(function(e) { return e.trim(); }).filter(Boolean);
    if (emails.length === 0) return;

    const name = subscriptionRow[COL.NAME];
    const amount = subscriptionRow[COL.AMOUNT];
    const currency = subscriptionRow[COL.CURRENCY];
    const nextDate = subscriptionRow[COL.NEXT_DATE];
    const payer = subscriptionRow[COL.PAYER] || '—';
    const payMethod = subscriptionRow[COL.PAY_METHOD] || '—';

    const daysText = daysUntil === 0 ? 'сегодня' : 'через ' + daysUntil + ' дн.';
    const subject = '💳 Напоминание: ' + name + ' — оплата ' + daysText;

    const body = 'Здравствуйте!\n\n' +
      'Напоминаем о предстоящей оплате:\n\n' +
      '📌 Подписка: ' + name + '\n' +
      '💰 Сумма: ' + amount + ' ' + currency + '\n' +
      '📅 Дата оплаты: ' + formatDate(nextDate) + '\n' +
      '👤 Кто платит: ' + payer + '\n' +
      '💳 Способ: ' + payMethod + '\n\n' +
      '⏰ Осталось дней: ' + daysUntil + '\n\n' +
      'Перейти к таблице: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

    for (const email of emails) {
      try {
        MailApp.sendEmail(email, subject, body);
      } catch (mailError) {
        console.error('Ошибка отправки на ' + email + ': ' + mailError.message);
      }
    }
  } catch (e) {
    console.error('sendReminder error: ' + e.message);
  }
}
