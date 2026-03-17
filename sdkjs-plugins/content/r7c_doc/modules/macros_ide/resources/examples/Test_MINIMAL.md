/**
 * МИНИМАЛЬНЫЙ ТЕСТ - просто вывести данные из MOCK_DATA.csv
 */

(function() {
    var sheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query('SELECT * FROM "MOCK_DATA.csv"', "MOCK_DATA.csv");

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    var df = result.result;

    // Заголовки
    for (var i = 0; i < df.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetValue(df.columns[i]);
    }

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        for (var c = 0; c < df.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(df.rows[r][c]);
        }
    }

    // Итого
    sheet.GetRange("A" + (df.rows.length + 3)).SetValue("Всего строк: " + df.rowCount);
})();
