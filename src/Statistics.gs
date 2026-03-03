// ==================== Statistics.gs ====================
// Обновление статистики на отдельном листе "📊 Статистика"
// Все суммы приводятся к валюте по умолчанию через курсы из Справочники!G:H

const MONTH_NAMES_RU = [
  'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
  'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
];

/**
 * Построить формулу суммы с конвертацией валют для листа Подписки.
 * Генерирует: SUMPRODUCT(cond1*cond2*P*rate_BYN*(E="BYN")) + SUMPRODUCT(...USD) + ...
 * Каждая валюта умножается на скалярную ссылку на ячейку курса — гарантированно работает.
 *
 * @param {string[]} conditions - массив условий вида '(Подписки!H$2:H$1000="Активна")'
 * @returns {string} формула
 */
function buildSubFormula_(conditions) {
  const conds = conditions.join('*');
  const parts = [];
  for (let i = 0; i < LOOKUP_CURRENCIES.length; i++) {
    const cur = LOOKUP_CURRENCIES[i];
    const rateCell = SHEET_LOOKUPS + '!$H$' + (i + 2);
    parts.push(
      'SUMPRODUCT(' + conds +
      '*(' + SHEET_SUBSCRIPTIONS + '!E$2:E$1000="' + cur + '")*' +
      '(' + SHEET_SUBSCRIPTIONS + '!P$2:P$1000))*' + rateCell
    );
  }
  return '=' + parts.join('+');
}

/**
 * Построить формулу суммы с конвертацией валют для листа История оплат.
 *
 * @param {string[]} conditions - массив условий
 * @returns {string} формула
 */
function buildHistFormula_(conditions) {
  const conds = conditions.join('*');
  const hist = "'" + SHEET_HISTORY + "'";
  const parts = [];
  for (let i = 0; i < LOOKUP_CURRENCIES.length; i++) {
    const cur = LOOKUP_CURRENCIES[i];
    const rateCell = SHEET_LOOKUPS + '!$H$' + (i + 2);
    parts.push(
      'SUMPRODUCT(' + conds +
      '*(' + hist + '!F$2:F$1000="' + cur + '")*' +
      '(' + hist + '!E$2:E$1000))*' + rateCell
    );
  }
  return '=' + parts.join('+');
}

/**
 * Обновить статистику — создаёт/обновляет лист "📊 Статистика"
 */
function updateStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateStatisticsSheet_(ss);

    const settings = getSettings();
    const defCurrency = settings['Валюта по умолчанию'] || 'BYN';

    sheet.clear();

    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(4, 200);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(3, 30);

    const SUB = SHEET_SUBSCRIPTIONS;
    const HIST = "'" + SHEET_HISTORY + "'";

    let row = 1;

    // === Секция 1: Сводка ===
    row = writeSectionHeader_(sheet, row, 'A', '📊 СВОДКА (' + defCurrency + ')');

    sheet.getRange(row, 1).setValue('Активных подписок:');
    sheet.getRange(row, 2).setFormula(
      '=COUNTIF(' + SUB + '!H:H;"Активна")'
    );
    row++;

    sheet.getRange(row, 1).setValue('Общая сумма/мес:');
    sheet.getRange(row, 2).setFormula(
      buildSubFormula_(['(' + SUB + '!H$2:H$1000="Активна")'])
    );
    sheet.getRange(row, 2).setNumberFormat('#,##0.00');
    row++;

    sheet.getRange(row, 1).setValue('Общая сумма/год:');
    sheet.getRange(row, 2).setFormula(
      '=(' + buildSubFormula_(['(' + SUB + '!H$2:H$1000="Активна")']).substring(1) + ')*12'
    );
    sheet.getRange(row, 2).setNumberFormat('#,##0.00');
    row++;

    sheet.getRange(row, 1).setValue('Оплачено за текущий месяц:');
    sheet.getRange(row, 2).setFormula(
      buildHistFormula_([
        '(' + HIST + '!D$2:D$1000>=DATE(YEAR(TODAY());MONTH(TODAY());1))'
      ])
    );
    sheet.getRange(row, 2).setNumberFormat('#,##0.00');
    row++;

    formatDataRows_(sheet, 2, row - 2, 'A');
    row++;

    // === Секция 2: По категориям ===
    row = writeSectionHeader_(sheet, row, 'A', '📂 ПО КАТЕГОРИЯМ (в мес., ' + defCurrency + ')');

    for (const category of LOOKUP_CATEGORIES) {
      sheet.getRange(row, 1).setValue(category);
      sheet.getRange(row, 2).setFormula(
        buildSubFormula_([
          '(' + SUB + '!H$2:H$1000="Активна")',
          '(' + SUB + '!C$2:C$1000="' + category + '")'
        ])
      );
      sheet.getRange(row, 2).setNumberFormat('#,##0.00');
      row++;
    }

    formatDataRows_(sheet, row - LOOKUP_CATEGORIES.length, LOOKUP_CATEGORIES.length, 'A');
    row++;

    // === Секция 3: По членам семьи ===
    row = writeSectionHeader_(sheet, row, 'A', '👥 ПО ЧЛЕНАМ СЕМЬИ (в мес., ' + defCurrency + ')');

    const subSheet = ss.getSheetByName(SHEET_SUBSCRIPTIONS);
    const data = subSheet.getDataRange().getValues();
    const payers = [];
    const payersSeen = {};
    for (let i = 1; i < data.length; i++) {
      const payer = data[i][COL.PAYER];
      if (payer && !payersSeen[payer]) {
        payers.push(payer);
        payersSeen[payer] = true;
      }
    }

    const payersStartRow = row;
    for (const payer of payers) {
      sheet.getRange(row, 1).setValue(payer);
      sheet.getRange(row, 2).setFormula(
        buildSubFormula_([
          '(' + SUB + '!H$2:H$1000="Активна")',
          '(' + SUB + '!M$2:M$1000="' + payer + '")'
        ])
      );
      sheet.getRange(row, 2).setNumberFormat('#,##0.00');
      row++;
    }

    if (payers.length > 0) {
      formatDataRows_(sheet, payersStartRow, payers.length, 'A');
    }

    // === Секция 4: По месяцам (колонки D-E) ===
    let rightRow = 1;
    rightRow = writeSectionHeader_(sheet, rightRow, 'D', '📅 ОПЛАТЫ ПО МЕСЯЦАМ (' + defCurrency + ')');

    const now = new Date();
    const monthsStartRow = rightRow;
    for (let i = 0; i < 12; i++) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const year = d.getFullYear();
      const month = d.getMonth();
      const nextMonth = month + 1;
      const nextYear = nextMonth > 11 ? year + 1 : year;
      const nextMonthAdj = nextMonth > 11 ? 0 : nextMonth;

      sheet.getRange(rightRow, 4).setValue(MONTH_NAMES_RU[month] + ' ' + year);
      sheet.getRange(rightRow, 5).setFormula(
        buildHistFormula_([
          '(' + HIST + '!D$2:D$1000>=DATE(' + year + ';' + (month + 1) + ';1))',
          '(' + HIST + '!D$2:D$1000<DATE(' + nextYear + ';' + (nextMonthAdj + 1) + ';1))'
        ])
      );
      sheet.getRange(rightRow, 5).setNumberFormat('#,##0.00');
      rightRow++;
    }

    formatDataRows_(sheet, monthsStartRow, 12, 'D');
    rightRow++;

    // === Секция 5: По годам (колонки D-E) ===
    rightRow = writeSectionHeader_(sheet, rightRow, 'D', '📆 ОПЛАТЫ ПО ГОДАМ (' + defCurrency + ')');

    const currentYear = now.getFullYear();
    const yearsStartRow = rightRow;
    for (let y = currentYear; y >= currentYear - 2; y--) {
      sheet.getRange(rightRow, 4).setValue(y);
      sheet.getRange(rightRow, 5).setFormula(
        buildHistFormula_([
          '(YEAR(' + HIST + '!D$2:D$1000)=' + y + ')'
        ])
      );
      sheet.getRange(rightRow, 5).setNumberFormat('#,##0.00');
      rightRow++;
    }

    formatDataRows_(sheet, yearsStartRow, 3, 'D');

    SpreadsheetApp.getActiveSpreadsheet().toast('📊 Статистика обновлена', 'Статистика', 3);
  } catch (e) {
    console.error('updateStatistics error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка обновления статистики: ' + e.message);
  }
}

/**
 * Получить или создать лист статистики
 */
function getOrCreateStatisticsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_STATISTICS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_STATISTICS);
  }
  return sheet;
}

/**
 * Написать заголовок секции (merge + стиль)
 */
function writeSectionHeader_(sheet, row, startCol, title) {
  const colNum = startCol === 'D' ? 4 : 1;
  const mergeRange = sheet.getRange(row, colNum, 1, 2);
  mergeRange.merge();
  mergeRange.setValue(title);
  mergeRange.setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#2C3E50')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left');
  return row + 1;
}

/**
 * Отформатировать строки данных (метки + значения)
 */
function formatDataRows_(sheet, startRow, count, startCol) {
  const colNum = startCol === 'D' ? 4 : 1;
  sheet.getRange(startRow, colNum, count, 1).setHorizontalAlignment('left');
  sheet.getRange(startRow, colNum + 1, count, 1).setHorizontalAlignment('right');
}
