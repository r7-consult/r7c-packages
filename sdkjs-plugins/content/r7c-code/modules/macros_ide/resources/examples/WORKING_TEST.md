/**
 * РАБОЧИЙ ТЕСТ - вывод всех данных из MOCK_DATA.csv
 * ВАЖНО: Используем правильное имя таблицы "MOCK_DATA.csv" (в кавычках)
 */

(function() {
    var sheet = Api.GetActiveSheet();

    // Очистка области
    sheet.GetRange("A1:Z100").Clear();

    // SQL запрос с правильным именем таблицы
    var result = Api.SqlNotebook.query(
        'SELECT * FROM "MOCK_DATA.csv"',
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ОШИБКА:");
        sheet.GetRange("A2").SetValue(result.error.code);
        sheet.GetRange("A3").SetValue(result.error.message);
        return;
    }

    var df = result.result;

    // Заголовки (жирным)
    for (var i = 0; i < df.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetValue(df.columns[i]);
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetBold(true);
    }

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        for (var c = 0; c < df.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(df.rows[r][c]);
        }
    }

    // Итого
    var totalRow = df.rows.length + 3;
    sheet.GetRange("A" + totalRow).SetValue("Всего строк:");
    sheet.GetRange("B" + totalRow).SetValue(df.rowCount);
    sheet.GetRange("A" + totalRow).SetBold(true);
    sheet.GetRange("B" + totalRow).SetBold(true);
})();
