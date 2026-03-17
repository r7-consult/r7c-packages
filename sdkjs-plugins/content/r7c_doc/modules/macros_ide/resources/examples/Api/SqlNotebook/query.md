/**
 * @file SqlNotebook_query_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SqlNotebook.query
 * @author R7-Consult
 * @version 1.0.0
 * @date February 2, 2026
 *
 * @description
 * This macro demonstrates how to execute a SQL query against a file loaded in SQLite Notebook
 * and export the results to an Excel spreadsheet. It queries data from a CSV file and formats
 * the output with headers, data rows, and a summary.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как выполнить SQL запрос к файлу, загруженному в SQLite Notebook,
 * и экспортировать результаты в таблицу Excel. Он запрашивает данные из CSV файла и форматирует
 * вывод с заголовками, строками данных и итогами.
 *
 * @param {string} sql - SQL query to execute (SQL запрос для выполнения)
 * @param {string} filename - Name of the loaded file (Имя загруженного файла)
 * @param {object} options - Optional parameters (Дополнительные параметры)
 *
 * @returns {object} Result object with ok status and result/error data
 *   {ok: true, result: {columns: [], rows: [], rowCount: n, columnCount: n}}
 *   OR {ok: false, error: {code: string, message: string}}
 *
 * @prerequisites
 * 1. Open SQLite Notebook tab in R7 Code plugin
 * 2. Load a file (e.g., MOCK_DATA.csv)
 * 3. Run this macro
 *
 * @see https://r7-consult.ru/docs/api/sqlnotebook
 */

(function() {
    'use strict';

    try {
        // Initialize R7 Office API
        // Инициализация API R7 Office
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }

        // Get active worksheet
        // Получение активного листа
        var oWorksheet = Api.GetActiveSheet();

        // ========== CONFIGURATION / КОНФИГУРАЦИЯ ==========

        // SQL Query - SELECT first 20 rows from MOCK_DATA.csv
        // SQL Запрос - выбрать первые 20 строк из MOCK_DATA.csv
        var sql = 'SELECT id, first_name, last_name, email, gender FROM "MOCK_DATA.csv" LIMIT 20';

        // Filename - must match the loaded file name in SQLite Notebook
        // Имя файла - должно совпадать с именем файла в SQLite Notebook
        var filename = 'MOCK_DATA.csv';

        // Options (optional) - autoAttach: automatically activate source if inactive
        // Опции (необязательно) - autoAttach: автоматически активировать источник
        var options = {
            autoAttach: true
        };

        // ========== EXECUTE QUERY / ВЫПОЛНЕНИЕ ЗАПРОСА ==========

        console.log('[SqlNotebook] Executing query:', sql);
        var result = Api.SqlNotebook.query(sql, filename, options);

        // ========== ERROR HANDLING / ОБРАБОТКА ОШИБОК ==========

        if (!result.ok) {
            // Display error in cell A1
            // Отобразить ошибку в ячейке A1
            oWorksheet.GetRange("A1").SetValue("ERROR: " + result.error.code);
            oWorksheet.GetRange("A2").SetValue(result.error.message);

            // Style error cells
            // Стиль ячеек с ошибкой
            var errorColor = Api.CreateColorFromRGB(255, 0, 0);
            oWorksheet.GetRange("A1").SetFontColor(errorColor);
            oWorksheet.GetRange("A1").SetBold(true);

            console.error('[SqlNotebook] Query failed:', result.error);
            return;
        }

        // ========== PROCESS RESULTS / ОБРАБОТКА РЕЗУЛЬТАТОВ ==========

        var data = result.result;
        console.log('[SqlNotebook] Query succeeded. Rows:', data.rowCount, 'Columns:', data.columnCount);

        var startRow = 1;

        // Title
        // Заголовок
        oWorksheet.GetRange("A" + startRow).SetValue("SQL Query Results");
        oWorksheet.GetRange("A" + startRow).SetBold(true);
        oWorksheet.GetRange("A" + startRow).SetFontSize(14);
        var titleColor = Api.CreateColorFromRGB(0, 100, 200);
        oWorksheet.GetRange("A" + startRow).SetFontColor(titleColor);
        startRow += 2;

        // Column Headers
        // Заголовки столбцов
        for (var col = 0; col < data.columns.length; col++) {
            var cell = oWorksheet.GetRange(String.fromCharCode(65 + col) + startRow);
            cell.SetValue(data.columns[col]);
            cell.SetBold(true);

            // Header background color
            // Цвет фона заголовков
            var headerBg = Api.CreateColorFromRGB(220, 230, 240);
            cell.SetFillColor(headerBg);
        }
        startRow++;

        // Data Rows
        // Строки данных
        for (var row = 0; row < data.rows.length; row++) {
            for (var col = 0; col < data.rows[row].length; col++) {
                var value = data.rows[row][col];
                var cellAddress = String.fromCharCode(65 + col) + (startRow + row);
                oWorksheet.GetRange(cellAddress).SetValue(value);
            }
        }
        startRow += data.rows.length + 1;

        // Summary
        // Итоги
        oWorksheet.GetRange("A" + startRow).SetValue("Total rows: " + data.rowCount);
        oWorksheet.GetRange("A" + startRow).SetBold(true);
        startRow++;
        oWorksheet.GetRange("A" + startRow).SetValue("Total columns: " + data.columnCount);
        oWorksheet.GetRange("A" + startRow).SetBold(true);
        startRow += 2;

        // Success message
        // Сообщение об успехе
        oWorksheet.GetRange("A" + startRow).SetValue("✅ Query executed successfully!");
        var successColor = Api.CreateColorFromRGB(0, 150, 0);
        oWorksheet.GetRange("A" + startRow).SetFontColor(successColor);

        console.log('[SqlNotebook] Macro executed successfully');

    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('[SqlNotebook] Macro execution failed:', error.message);

        // Display error to user
        // Отобразить ошибку пользователю
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
                var errorColor = Api.CreateColorFromRGB(255, 0, 0);
                sheet.GetRange('A1').SetFontColor(errorColor);
                sheet.GetRange('A1').SetBold(true);
            }
        }
    }
})();
