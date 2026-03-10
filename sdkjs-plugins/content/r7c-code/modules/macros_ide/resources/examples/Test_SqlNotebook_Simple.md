/**
 * Простой тест Api.SqlNotebook.query()
 * Просто выводит все данные из MOCK_DATA.csv в таблицу
 */

(function() {
    'use strict';

    try {
        // Получаем активный лист
        var sheet = Api.GetActiveSheet();

        // Очищаем область
        sheet.GetRange("A1:Z100").Clear();

        // Выполняем SQL запрос
        var result = Api.SqlNotebook.query(
            'SELECT * FROM "MOCK_DATA.csv"',
            "MOCK_DATA.csv"
        );

        // Проверка на ошибки
        if (!result.ok) {
            sheet.GetRange("A1").SetValue("ОШИБКА:");
            sheet.GetRange("A2").SetValue(result.error.code);
            sheet.GetRange("A3").SetValue(result.error.message);
            return;
        }

        var df = result.result;

        // Выводим заголовки (жирным шрифтом)
        for (var i = 0; i < df.columns.length; i++) {
            var headerCell = String.fromCharCode(65 + i) + "1";
            sheet.GetRange(headerCell).SetValue(df.columns[i]);
            sheet.GetRange(headerCell).SetBold(true);
        }

        // Выводим данные
        for (var row = 0; row < df.rows.length; row++) {
            for (var col = 0; col < df.rows[row].length; col++) {
                var cellAddr = String.fromCharCode(65 + col) + (row + 2);
                sheet.GetRange(cellAddr).SetValue(df.rows[row][col]);
            }
        }

        // Итоговая информация внизу
        var summaryRow = df.rows.length + 3;
        sheet.GetRange("A" + summaryRow).SetValue("Загружено строк:");
        sheet.GetRange("B" + summaryRow).SetValue(df.rowCount);
        sheet.GetRange("A" + summaryRow).SetBold(true);

    } catch (error) {
        var sheet = Api.GetActiveSheet();
        sheet.GetRange("A1").SetValue("ОШИБКА ВЫПОЛНЕНИЯ:");
        sheet.GetRange("A2").SetValue(error.message);
        console.error(error);
    }
})();
