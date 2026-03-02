// ==================== Statistics.gs ====================
// Обновление статистики и дашборда

/**
 * Обновить статистику на листе "Подписки" (область ниже данных)
 */
function updateStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SUBSCRIPTIONS);
    const data = sheet.getDataRange().getValues();

    // Найти последнюю строку с данными
    let lastDataRow = 1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.NAME]) lastDataRow = i + 1;
    }

    // Начало блока статистики — через 3 строки после данных
    const startRow = Math.max(lastDataRow + 3, 50);

    // Очистить область статистики
    sheet.getRange(startRow, 1, 20, 4).clear();

    // Заголовок
    sheet.getRange(startRow, 1).setValue('📊 СТАТИСТИКА');
    sheet.getRange(startRow, 1).setFontWeight('bold').setFontSize(12);

    let row = startRow + 1;

    // Общая сумма/мес
    sheet.getRange(row, 1).setValue('Общая сумма/мес:');
    sheet.getRange(row, 2).setFormula(
      '=SUMIF(H2:H' + lastDataRow + ',"Активна",P2:P' + lastDataRow + ')'
    );
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;

    // Общая сумма/год
    sheet.getRange(row, 1).setValue('Общая сумма/год:');
    sheet.getRange(row, 2).setFormula(
      '=SUMIF(H2:H' + lastDataRow + ',"Активна",P2:P' + lastDataRow + ')*12'
    );
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;

    // Всего активных подписок
    sheet.getRange(row, 1).setValue('Активных подписок:');
    sheet.getRange(row, 2).setFormula(
      '=COUNTIF(H2:H' + lastDataRow + ',"Активна")'
    );
    row++;

    // Оплачено за текущий месяц (из Истории оплат)
    sheet.getRange(row, 1).setValue('Оплачено за месяц:');
    sheet.getRange(row, 2).setFormula(
      '=SUMPRODUCT((\'' + SHEET_HISTORY + '\'!D2:D1000>=DATE(YEAR(TODAY()),MONTH(TODAY()),1))*' +
      '(\'' + SHEET_HISTORY + '\'!E2:E1000))'
    );
    row += 2;

    // По категориям
    sheet.getRange(row, 1).setValue('📂 По категориям (в мес.):');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;

    for (const category of LOOKUP_CATEGORIES) {
      sheet.getRange(row, 1).setValue('  ' + category + ':');
      sheet.getRange(row, 2).setFormula(
        '=SUMIFS(P2:P' + lastDataRow + ',H2:H' + lastDataRow + ',"Активна",C2:C' + lastDataRow + ',"' + category + '")'
      );
      row++;
    }

    row++;

    // По членам семьи
    sheet.getRange(row, 1).setValue('👥 По членам семьи (в мес.):');
    sheet.getRange(row, 1).setFontWeight('bold');
    row++;

    // Собрать уникальных плательщиков из данных
    const payers = new Set();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL.PAYER]) payers.add(data[i][COL.PAYER]);
    }

    for (const payer of payers) {
      sheet.getRange(row, 1).setValue('  ' + payer + ':');
      sheet.getRange(row, 2).setFormula(
        '=SUMIFS(P2:P' + lastDataRow + ',H2:H' + lastDataRow + ',"Активна",M2:M' + lastDataRow + ',"' + payer + '")'
      );
      row++;
    }

    SpreadsheetApp.getActiveSpreadsheet().toast('📊 Статистика обновлена', 'Статистика', 3);
  } catch (e) {
    console.error('updateStatistics error: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ Ошибка обновления статистики: ' + e.message);
  }
}
