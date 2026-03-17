/**
 * @file SqlNotebook_batch_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SqlNotebook.batch
 * @author R7-Consult
 * @version 1.0.0
 * @date February 2, 2026
 *
 * @description
 * This macro demonstrates how to execute multiple SQL queries at once using batch operations.
 * It runs several queries against the same or different data sources and presents all results
 * in a formatted spreadsheet with individual query results and summary statistics.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как выполнить несколько SQL запросов одновременно используя пакетные операции.
 * Он выполняет несколько запросов к одному или разным источникам данных и представляет все результаты
 * в отформатированной таблице с индивидуальными результатами запросов и итоговой статистикой.
 *
 * @param {Array} queries - Array of query objects [{sql, filename, options}]
 *
 * @returns {object} Result object with ok status and array of results
 *   {ok: true, results: [{ok: true, result: {...}}, {ok: false, error: {...}}, ...]}
 *
 * @prerequisites
 * 1. Open SQLite Notebook tab in R7 Code plugin
 * 2. Load one or more files (e.g., MOCK_DATA.csv)
 * 3. Run this macro
 *
 * @see https://r7-consult.ru/docs/api/sqlnotebook/batch
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

        // Define multiple queries to execute
        // Определение нескольких запросов для выполнения
        var queries = [
            // Query 1: Total record count
            // Запрос 1: Общее количество записей
            {
                sql: 'SELECT COUNT(*) as total FROM "MOCK_DATA.csv"',
                filename: 'MOCK_DATA.csv'
            },
            // Query 2: Gender distribution
            // Запрос 2: Распределение по полу
            {
                sql: 'SELECT gender, COUNT(*) as count FROM "MOCK_DATA.csv" GROUP BY gender',
                filename: 'MOCK_DATA.csv'
            },
            // Query 3: Top 10 records by ID
            // Запрос 3: Топ 10 записей по ID
            {
                sql: 'SELECT id, first_name, last_name, email FROM "MOCK_DATA.csv" ORDER BY id DESC LIMIT 10',
                filename: 'MOCK_DATA.csv'
            },
            // Query 4: Email domains analysis
            // Запрос 4: Анализ доменов email
            {
                sql: 'SELECT SUBSTR(email, INSTR(email, "@") + 1) as domain, COUNT(*) as count FROM "MOCK_DATA.csv" GROUP BY domain LIMIT 10',
                filename: 'MOCK_DATA.csv'
            }
        ];

        // ========== EXECUTE BATCH / ВЫПОЛНЕНИЕ ПАКЕТА ==========

        console.log('[SqlNotebook] Executing batch with', queries.length, 'queries');
        var batchResult = Api.SqlNotebook.batch(queries);

        // ========== ERROR HANDLING / ОБРАБОТКА ОШИБОК ==========

        if (!batchResult.ok) {
            // Display batch-level error
            // Отобразить ошибку уровня пакета
            oWorksheet.GetRange("A1").SetValue("BATCH ERROR: " + batchResult.error.code);
            oWorksheet.GetRange("A2").SetValue(batchResult.error.message);

            var errorColor = Api.CreateColorFromRGB(255, 0, 0);
            oWorksheet.GetRange("A1").SetFontColor(errorColor);
            oWorksheet.GetRange("A1").SetBold(true);

            console.error('[SqlNotebook] Batch failed:', batchResult.error);
            return;
        }

        // ========== PROCESS RESULTS / ОБРАБОТКА РЕЗУЛЬТАТОВ ==========

        var results = batchResult.results;
        console.log('[SqlNotebook] Batch succeeded. Processing', results.length, 'results');

        var currentRow = 1;

        // Main Title
        // Главный заголовок
        oWorksheet.GetRange("A" + currentRow).SetValue("Batch Query Results");
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        oWorksheet.GetRange("A" + currentRow).SetFontSize(16);
        var titleColor = Api.CreateColorFromRGB(0, 80, 160);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(titleColor);
        currentRow += 2;

        // Batch summary
        // Итоги пакета
        oWorksheet.GetRange("A" + currentRow).SetValue("Total queries: " + results.length);
        currentRow++;

        var successCount = 0;
        var failureCount = 0;
        results.forEach(function(result) {
            if (result.ok) successCount++;
            else failureCount++;
        });

        oWorksheet.GetRange("A" + currentRow).SetValue("Successful: " + successCount);
        var successColor = Api.CreateColorFromRGB(0, 150, 0);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(successColor);
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("Failed: " + failureCount);
        if (failureCount > 0) {
            var failColor = Api.CreateColorFromRGB(255, 0, 0);
            oWorksheet.GetRange("A" + currentRow).SetFontColor(failColor);
        }
        currentRow += 2;

        // ========== DISPLAY EACH QUERY RESULT / ОТОБРАЖЕНИЕ КАЖДОГО РЕЗУЛЬТАТА ==========

        results.forEach(function(result, index) {
            // Query header
            // Заголовок запроса
            oWorksheet.GetRange("A" + currentRow).SetValue("Query " + (index + 1) + ":");
            oWorksheet.GetRange("A" + currentRow).SetBold(true);
            oWorksheet.GetRange("A" + currentRow).SetFontSize(12);

            var queryHeaderColor = Api.CreateColorFromRGB(50, 120, 180);
            oWorksheet.GetRange("A" + currentRow).SetFontColor(queryHeaderColor);
            currentRow++;

            // Check if query succeeded
            // Проверка успешности запроса
            if (!result.ok) {
                // Display error for this query
                // Отобразить ошибку для этого запроса
                oWorksheet.GetRange("A" + currentRow).SetValue("ERROR: " + result.error.message);
                var errorColor = Api.CreateColorFromRGB(255, 0, 0);
                oWorksheet.GetRange("A" + currentRow).SetFontColor(errorColor);
                currentRow += 2;
                return;
            }

            var data = result.result;

            // Column headers
            // Заголовки столбцов
            for (var col = 0; col < data.columns.length; col++) {
                var cell = oWorksheet.GetRange(String.fromCharCode(65 + col) + currentRow);
                cell.SetValue(data.columns[col]);
                cell.SetBold(true);

                var headerBg = Api.CreateColorFromRGB(230, 240, 250);
                cell.SetFillColor(headerBg);
            }
            currentRow++;

            // Data rows (limit to 10 rows per query for readability)
            // Строки данных (максимум 10 строк на запрос для читаемости)
            var maxRows = Math.min(data.rows.length, 10);
            for (var row = 0; row < maxRows; row++) {
                for (var col = 0; col < data.rows[row].length; col++) {
                    var value = data.rows[row][col];
                    var cellAddress = String.fromCharCode(65 + col) + currentRow;
                    oWorksheet.GetRange(cellAddress).SetValue(value);
                }
                currentRow++;
            }

            // Show truncation notice if needed
            // Показать уведомление об усечении, если нужно
            if (data.rows.length > 10) {
                oWorksheet.GetRange("A" + currentRow).SetValue("... and " + (data.rows.length - 10) + " more rows");
                oWorksheet.GetRange("A" + currentRow).SetItalic(true);
                currentRow++;
            }

            // Query summary
            // Итоги запроса
            oWorksheet.GetRange("A" + currentRow).SetValue("Rows returned: " + data.rowCount);
            oWorksheet.GetRange("A" + currentRow).SetItalic(true);
            currentRow += 2;
        });

        // ========== FINAL MESSAGE / ФИНАЛЬНОЕ СООБЩЕНИЕ ==========

        currentRow++;
        oWorksheet.GetRange("A" + currentRow).SetValue("✅ Batch execution completed!");
        var doneColor = Api.CreateColorFromRGB(0, 150, 0);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(doneColor);
        oWorksheet.GetRange("A" + currentRow).SetBold(true);

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
