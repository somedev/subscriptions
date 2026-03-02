// ==================== Utils.gs ====================
// Вспомогательные функции

/**
 * Безопасное добавление месяцев к дате (обработка конца месяца)
 * Jan 31 + 1 month = Feb 28 (или 29 в високосный год)
 */
function addMonths(date, months) {
  const result = new Date(date);
  const day = result.getDate();
  result.setMonth(result.getMonth() + months);
  if (result.getDate() !== day) {
    result.setDate(0); // Последний день предыдущего месяца
  }
  return result;
}

/**
 * Вычисление следующей даты оплаты по периоду
 */
function calculateNextDate(currentDate, period) {
  const date = new Date(currentDate);
  switch (period) {
    case 'Неделя':
      date.setDate(date.getDate() + 7);
      return date;
    case 'Месяц':
      return addMonths(date, 1);
    case 'Квартал':
      return addMonths(date, 3);
    case 'Полгода':
      return addMonths(date, 6);
    case 'Год':
      return addMonths(date, 12);
    default: {
      // Произвольный период: "N мес." (например, "2 мес.", "5 мес.")
      const m = period.match(/^(\d+)\s*мес\.?$/);
      if (m) return addMonths(date, parseInt(m[1]));
      return addMonths(date, 1);
    }
  }
}

/**
 * Получить лист по имени с проверкой существования
 */
function getSheetByName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Лист "' + name + '" не найден. Запустите первоначальную настройку.');
  }
  return sheet;
}

/**
 * Чтение настроек из листа "Настройки"
 */
function getSettings() {
  const sheet = getSheetByName(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]) {
      settings[data[i][0]] = data[i][1];
    }
  }
  return settings;
}

/**
 * Форматирование даты в русском формате
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) return '';
  return Utilities.formatDate(date, 'Europe/Minsk', 'dd.MM.yyyy');
}

/**
 * Получить следующий ID для автоинкремента
 */
function getNextId(sheet) {
  const data = sheet.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < data.length; i++) {
    const id = parseInt(data[i][0]);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }
  return maxId + 1;
}

/**
 * Получить или создать календарь по имени
 */
function getOrCreateCalendar(calendarName) {
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  }
  return CalendarApp.createCalendar(calendarName);
}
