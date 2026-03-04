// ==================== WebApi.gs ====================
// Серверные API-функции для мобильного веб-приложения

/**
 * Получить ВСЕ данные за один вызов — подписки, статистику, историю, настройки.
 * Минимизирует количество google.script.run round-trips (1-3 сек каждый).
 */
function getAllDataApi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subSheet = ss.getSheetByName(SHEET_SUBSCRIPTIONS);
  const subData = subSheet ? subSheet.getDataRange().getValues() : [];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // --- Подписки ---
  const subscriptions = [];
  let totalCount = 0;
  let nextPayment = null;
  const byCategory = {};

  for (let i = 1; i < subData.length; i++) {
    const row = subData[i];
    if (!row[COL.NAME]) continue;

    const nextDate = row[COL.NEXT_DATE] ? new Date(row[COL.NEXT_DATE]) : null;
    const lastPaid = row[COL.LAST_PAID] ? new Date(row[COL.LAST_PAID]) : null;
    let daysUntil = null;
    if (nextDate) {
      const nd = new Date(nextDate);
      nd.setHours(0, 0, 0, 0);
      daysUntil = Math.round((nd - today) / (1000 * 60 * 60 * 24));
    }

    const sub = {
      rowIndex: i + 1,
      id: row[COL.ID],
      name: row[COL.NAME],
      category: row[COL.CATEGORY],
      amount: row[COL.AMOUNT],
      currency: row[COL.CURRENCY],
      period: row[COL.PERIOD],
      nextDate: nextDate ? nextDate.toISOString() : null,
      status: row[COL.STATUS],
      lastPaid: lastPaid ? lastPaid.toISOString() : null,
      notify: row[COL.NOTIFICATIONS],
      remindDays: row[COL.REMIND_DAYS],
      payer: row[COL.PAYER],
      payMethod: row[COL.PAY_METHOD],
      notes: row[COL.NOTES],
      daysUntil: daysUntil
    };
    subscriptions.push(sub);

    // Статистика (считаем параллельно, пока итерируем)
    if (row[COL.STATUS] === 'Активна') {
      totalCount++;
      const amount = parseFloat(row[COL.AMOUNT]) || 0;
      const period = row[COL.PERIOD];
      let monthly = 0;
      switch (period) {
        case 'Неделя': monthly = amount * 4.33; break;
        case 'Месяц': monthly = amount; break;
        case 'Квартал': monthly = amount / 3; break;
        case 'Полгода': monthly = amount / 6; break;
        case 'Год': monthly = amount / 12; break;
        default: {
          const m = period.match(/^(\d+)\s*мес\.?$/);
          monthly = m ? amount / parseInt(m[1]) : amount;
        }
      }
      const currency = row[COL.CURRENCY] || '';
      const cat = row[COL.CATEGORY] || 'Другое';
      if (!byCategory[cat]) byCategory[cat] = {};
      if (!byCategory[cat][currency]) byCategory[cat][currency] = 0;
      byCategory[cat][currency] += monthly;

      if (nextDate) {
        if (!nextPayment || nextDate < new Date(nextPayment.date)) {
          nextPayment = { name: row[COL.NAME], date: nextDate.toISOString(), amount: amount, currency: currency };
        }
      }
    }
  }

  subscriptions.sort(function(a, b) {
    if (a.status !== 'Активна' && b.status === 'Активна') return 1;
    if (a.status === 'Активна' && b.status !== 'Активна') return -1;
    if (a.daysUntil === null && b.daysUntil === null) return 0;
    if (a.daysUntil === null) return 1;
    if (b.daysUntil === null) return -1;
    return a.daysUntil - b.daysUntil;
  });

  // Категории в массив
  const categoryList = [];
  for (const cat in byCategory) {
    const amounts = [];
    for (const cur in byCategory[cat]) {
      amounts.push({ currency: cur, monthly: Math.round(byCategory[cat][cur] * 100) / 100 });
    }
    categoryList.push({ category: cat, amounts: amounts });
  }
  categoryList.sort(function(a, b) {
    const tA = a.amounts.reduce(function(s, x) { return s + x.monthly; }, 0);
    const tB = b.amounts.reduce(function(s, x) { return s + x.monthly; }, 0);
    return tB - tA;
  });

  // --- История ---
  let history = [];
  try {
    const histSheet = ss.getSheetByName(SHEET_HISTORY);
    if (histSheet) {
      const histData = histSheet.getDataRange().getValues();
      const start = Math.max(1, histData.length - 10);
      for (let i = histData.length - 1; i >= start; i--) {
        const row = histData[i];
        if (!row[2]) continue;
        history.push({
          date: row[3] ? new Date(row[3]).toISOString() : null,
          name: row[2],
          amount: row[4],
          currency: row[5]
        });
      }
    }
  } catch (e) { /* нет листа истории */ }

  // --- Настройки ---
  const settings = getSettings();
  const familyStr = settings['Члены семьи'] || '';
  const family = familyStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
  let sheetUrl = '';
  try { sheetUrl = ss.getUrl(); } catch (e) { /* web app context */ }

  return {
    subscriptions: subscriptions,
    stats: { totalCount: totalCount, nextPayment: nextPayment, byCategory: categoryList },
    history: history,
    settings: {
      categories: LOOKUP_CATEGORIES,
      periods: LOOKUP_PERIODS,
      currencies: LOOKUP_CURRENCIES,
      payMethods: LOOKUP_PAY_METHODS,
      family: family,
      defaultCurrency: settings['Валюта по умолчанию'] || 'BYN',
      email: settings['Email для уведомлений'] || '',
      reminderDays: parseInt(settings['Дней до напоминания']) || 3,
      calendarName: settings['Название календаря'] || '💳 Подписки',
      sheetUrl: sheetUrl
    }
  };
}

/**
 * Получить список подписок, отсортированный по срочности
 */
function getSubscriptionsApi() {
  const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const subscriptions = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.NAME]) continue;

    const nextDate = row[COL.NEXT_DATE] ? new Date(row[COL.NEXT_DATE]) : null;
    const lastPaid = row[COL.LAST_PAID] ? new Date(row[COL.LAST_PAID]) : null;

    let daysUntil = null;
    if (nextDate) {
      const nd = new Date(nextDate);
      nd.setHours(0, 0, 0, 0);
      daysUntil = Math.round((nd - today) / (1000 * 60 * 60 * 24));
    }

    subscriptions.push({
      rowIndex: i + 1,
      id: row[COL.ID],
      name: row[COL.NAME],
      category: row[COL.CATEGORY],
      amount: row[COL.AMOUNT],
      currency: row[COL.CURRENCY],
      period: row[COL.PERIOD],
      nextDate: nextDate ? nextDate.toISOString() : null,
      status: row[COL.STATUS],
      lastPaid: lastPaid ? lastPaid.toISOString() : null,
      notify: row[COL.NOTIFICATIONS],
      remindDays: row[COL.REMIND_DAYS],
      payer: row[COL.PAYER],
      payMethod: row[COL.PAY_METHOD],
      notes: row[COL.NOTES],
      daysUntil: daysUntil
    });
  }

  // Сортировка: просроченные первые, потом скоро, потом остальные
  subscriptions.sort(function(a, b) {
    // Неактивные — в конец
    if (a.status !== 'Активна' && b.status === 'Активна') return 1;
    if (a.status === 'Активна' && b.status !== 'Активна') return -1;
    // По дням до оплаты
    if (a.daysUntil === null && b.daysUntil === null) return 0;
    if (a.daysUntil === null) return 1;
    if (b.daysUntil === null) return -1;
    return a.daysUntil - b.daysUntil;
  });

  return subscriptions;
}

/**
 * Добавить подписку через веб-форму
 */
function addSubscriptionApi(data) {
  const result = processAddForm({
    name: data.name,
    category: data.category,
    amount: data.amount,
    currency: data.currency,
    period: data.period,
    nextDate: data.nextDate,
    status: 'Активна',
    notifications: data.notify !== false,
    remindDays: data.remindDays || 3,
    payer: data.payer || '',
    payMethod: data.payMethod || '',
    notes: data.notes || ''
  });

  if (!result.success) {
    throw new Error(result.error);
  }

  // Синхронизировать календарь для новой подписки
  try { syncCalendar(); } catch (e) { console.error('Calendar sync error: ' + e.message); }

  return result;
}

/**
 * Обновить существующую подписку
 */
function updateSubscriptionApi(rowIndex, data) {
  try {
    const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
    const range = sheet.getRange(rowIndex, 1, 1, HEADERS_SUBSCRIPTIONS.length);
    const row = range.getValues()[0];

    if (!row[COL.NAME]) {
      throw new Error('Строка не найдена');
    }

    // Обновить поля
    sheet.getRange(rowIndex, COL.NAME + 1).setValue(data.name);
    sheet.getRange(rowIndex, COL.CATEGORY + 1).setValue(data.category);
    sheet.getRange(rowIndex, COL.AMOUNT + 1).setValue(parseFloat(data.amount));
    sheet.getRange(rowIndex, COL.CURRENCY + 1).setValue(data.currency);
    sheet.getRange(rowIndex, COL.PERIOD + 1).setValue(data.period);
    sheet.getRange(rowIndex, COL.NEXT_DATE + 1).setValue(new Date(data.nextDate));
    sheet.getRange(rowIndex, COL.NOTIFICATIONS + 1).setValue(data.notify !== false);
    sheet.getRange(rowIndex, COL.REMIND_DAYS + 1).setValue(parseInt(data.remindDays) || 3);
    sheet.getRange(rowIndex, COL.PAYER + 1).setValue(data.payer || '');
    sheet.getRange(rowIndex, COL.PAY_METHOD + 1).setValue(data.payMethod || '');
    sheet.getRange(rowIndex, COL.NOTES + 1).setValue(data.notes || '');

    // Обновить формулы
    sheet.getRange(rowIndex, COL.MONTHLY_COST + 1).setFormula(
      '=IF(H' + rowIndex + '="Активна";' +
      'SWITCH(F' + rowIndex + ';"Месяц";D' + rowIndex + ';"Квартал";D' + rowIndex + '/3;' +
      '"Полгода";D' + rowIndex + '/6;"Год";D' + rowIndex + '/12;"Неделя";D' + rowIndex + '*4,33;' +
      'IFERROR(D' + rowIndex + '/VALUE(LEFT(F' + rowIndex + ';LEN(F' + rowIndex + ')-5));0));0)'
    );
    sheet.getRange(rowIndex, COL.DAYS_UNTIL + 1).setFormula(
      '=IF(AND(H' + rowIndex + '="Активна"; G' + rowIndex + '<>""); G' + rowIndex + '-TODAY(); "")'
    );

    // Синхронизировать календарь
    try { syncCalendar(); } catch (e) { console.error('Calendar sync error: ' + e.message); }

    return { success: true };
  } catch (e) {
    console.error('updateSubscriptionApi error: ' + e.message);
    throw e;
  }
}

/**
 * Удалить подписку
 */
function deleteSubscriptionApi(rowIndex) {
  try {
    const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
    const row = sheet.getRange(rowIndex, 1, 1, HEADERS_SUBSCRIPTIONS.length).getValues()[0];

    // Удалить событие из календаря
    const eventId = row[COL.CALENDAR_ID];
    if (eventId) {
      try {
        const settings = getSettings();
        const calendarName = settings['Название календаря'] || '💳 Подписки';
        const calendars = CalendarApp.getCalendarsByName(calendarName);
        if (calendars.length > 0) {
          const event = calendars[0].getEventById(eventId);
          if (event) event.deleteEvent();
        }
      } catch (e) { console.error('Calendar delete error: ' + e.message); }
    }

    // Удалить строку
    sheet.deleteRow(rowIndex);

    return { success: true };
  } catch (e) {
    console.error('deleteSubscriptionApi error: ' + e.message);
    throw e;
  }
}

/**
 * Отметить оплату — переиспользует существующую логику confirmPayment
 */
function markAsPaidApi(rowIndex) {
  confirmPayment(rowIndex);

  // Вернуть обновлённые данные строки
  const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
  const row = sheet.getRange(rowIndex, 1, 1, HEADERS_SUBSCRIPTIONS.length).getValues()[0];
  const nextDate = row[COL.NEXT_DATE] ? new Date(row[COL.NEXT_DATE]) : null;

  return {
    rowIndex: rowIndex,
    name: row[COL.NAME],
    nextDate: nextDate ? nextDate.toISOString() : null,
    lastPaid: new Date().toISOString()
  };
}

/**
 * Получить статистику для веб-интерфейса
 */
function getStatsApi() {
  const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let totalCount = 0;
  let monthlyTotal = 0;
  let nextPayment = null;
  const byCategory = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.NAME] || row[COL.STATUS] !== 'Активна') continue;

    totalCount++;

    // Расчёт месячной стоимости
    const amount = parseFloat(row[COL.AMOUNT]) || 0;
    const period = row[COL.PERIOD];
    let monthly = 0;
    switch (period) {
      case 'Неделя': monthly = amount * 4.33; break;
      case 'Месяц': monthly = amount; break;
      case 'Квартал': monthly = amount / 3; break;
      case 'Полгода': monthly = amount / 6; break;
      case 'Год': monthly = amount / 12; break;
      default: {
        const m = period.match(/^(\d+)\s*мес\.?$/);
        monthly = m ? amount / parseInt(m[1]) : amount;
      }
    }

    const currency = row[COL.CURRENCY] || '';
    const cat = row[COL.CATEGORY] || 'Другое';

    if (!byCategory[cat]) byCategory[cat] = {};
    if (!byCategory[cat][currency]) byCategory[cat][currency] = 0;
    byCategory[cat][currency] += monthly;

    // Ближайшая оплата
    const nextDate = row[COL.NEXT_DATE] ? new Date(row[COL.NEXT_DATE]) : null;
    if (nextDate) {
      if (!nextPayment || nextDate < new Date(nextPayment.date)) {
        nextPayment = {
          name: row[COL.NAME],
          date: nextDate.toISOString(),
          amount: amount,
          currency: currency
        };
      }
    }
  }

  // Агрегация по категориям в массив
  const categoryList = [];
  for (const cat in byCategory) {
    const amounts = [];
    for (const cur in byCategory[cat]) {
      amounts.push({ currency: cur, monthly: Math.round(byCategory[cat][cur] * 100) / 100 });
    }
    categoryList.push({ category: cat, amounts: amounts });
  }
  categoryList.sort(function(a, b) {
    const totalA = a.amounts.reduce(function(s, x) { return s + x.monthly; }, 0);
    const totalB = b.amounts.reduce(function(s, x) { return s + x.monthly; }, 0);
    return totalB - totalA;
  });

  return {
    totalCount: totalCount,
    nextPayment: nextPayment,
    byCategory: categoryList
  };
}

/**
 * Получить историю оплат
 */
function getHistoryApi(limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_HISTORY);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    const history = [];

    // Последние записи (от конца к началу)
    const start = Math.max(1, data.length - (limit || 10));
    for (let i = data.length - 1; i >= start; i--) {
      const row = data[i];
      if (!row[2]) continue; // Нет названия
      history.push({
        date: row[3] ? new Date(row[3]).toISOString() : null,
        name: row[2],
        amount: row[4],
        currency: row[5]
      });
    }

    return history;
  } catch (e) {
    console.error('getHistoryApi error: ' + e.message);
    return [];
  }
}

/**
 * Получить настройки для веб-интерфейса
 */
function getSettingsApi() {
  const settings = getSettings();
  const familyStr = settings['Члены семьи'] || '';
  const family = familyStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);

  let sheetUrl = '';
  try { sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl(); } catch (e) { /* web app context */ }

  return {
    categories: LOOKUP_CATEGORIES,
    periods: LOOKUP_PERIODS,
    currencies: LOOKUP_CURRENCIES,
    payMethods: LOOKUP_PAY_METHODS,
    family: family,
    defaultCurrency: settings['Валюта по умолчанию'] || 'BYN',
    email: settings['Email для уведомлений'] || '',
    reminderDays: parseInt(settings['Дней до напоминания']) || 3,
    calendarName: settings['Название календаря'] || '💳 Подписки',
    sheetUrl: sheetUrl
  };
}

/**
 * Получить одну подписку по rowIndex (для формы редактирования)
 */
function getSubscriptionApi(rowIndex) {
  const sheet = getSheetByName(SHEET_SUBSCRIPTIONS);
  const row = sheet.getRange(rowIndex, 1, 1, HEADERS_SUBSCRIPTIONS.length).getValues()[0];

  if (!row[COL.NAME]) throw new Error('Подписка не найдена');

  const nextDate = row[COL.NEXT_DATE] ? new Date(row[COL.NEXT_DATE]) : null;
  const lastPaid = row[COL.LAST_PAID] ? new Date(row[COL.LAST_PAID]) : null;

  return {
    rowIndex: rowIndex,
    id: row[COL.ID],
    name: row[COL.NAME],
    category: row[COL.CATEGORY],
    amount: row[COL.AMOUNT],
    currency: row[COL.CURRENCY],
    period: row[COL.PERIOD],
    nextDate: nextDate ? nextDate.toISOString() : null,
    status: row[COL.STATUS],
    lastPaid: lastPaid ? lastPaid.toISOString() : null,
    notify: row[COL.NOTIFICATIONS],
    remindDays: row[COL.REMIND_DAYS],
    payer: row[COL.PAYER],
    payMethod: row[COL.PAY_METHOD],
    notes: row[COL.NOTES]
  };
}
