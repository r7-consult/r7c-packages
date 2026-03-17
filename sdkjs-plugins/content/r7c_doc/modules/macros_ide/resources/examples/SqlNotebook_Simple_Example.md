/**
 * @file SqlNotebook_Simple_Example_macroR7.js
 * @brief R7 Office JavaScript Macro - Simple SQLite Notebook Query Example
 * @author R7-Consult
 * @version 1.0.0
 * @date February 2, 2026
 *
 * @description
 * A simple example showing how to query a CSV file loaded in SQLite Notebook
 * and display the results in an Excel sheet.
 *
 * @description (Russian)
 * Простой пример показывает как выполнить запрос к CSV файлу, загруженному
 * в SQLite Notebook, и отобразить результаты в таблице Excel.
 */

(function() {
    'use strict';

    try {
        // Выполнить SQL запрос к загруженному файлу
        var result = Api.SqlNotebook.query(
            "SELECT * FROM data WHERE value > 100 LIMIT 10",
            "MOCK_DATA.csv"
        );

        // Проверка на ошибки
        if (!result.ok) {
            Api.ShowMessage("Ошибка", result.error.message);
            return;
        }

        // Получение активного листа
        var sheet = Api.GetActiveSheet();
        var df = result.result;

        // Вывод заголовков
        df.columns.forEach(function(col, i) {
            var cell = String.fromCharCode(65 + i) + '1';
            sheet.GetRange(cell).SetValue(col);
        });

        // Вывод данных
        df.rows.forEach(function(row, rowIdx) {
            row.forEach(function(val, colIdx) {
                var cell = String.fromCharCode(65 + colIdx) + (rowIdx + 2);
                sheet.GetRange(cell).SetValue(val);
            });
        });

        // Уведомление об успехе
        Api.ShowMessage("Успех", "Загружено " + df.rowCount + " строк");

    } catch (error) {
        Api.ShowMessage('Ошибка', error.message);
    }
})();
