/**
 * @file SqlNotebook_getSources_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SqlNotebook.getSources
 * @author R7-Consult
 * @version 1.0.0
 * @date February 2, 2026
 *
 * @description
 * This macro demonstrates how to retrieve the list of all data sources currently loaded
 * in SQLite Notebook. It displays source information including name, format, active status,
 * and provides example queries for each available source.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить список всех источников данных, загруженных
 * в SQLite Notebook. Он отображает информацию об источниках, включая имя, формат, статус активности
 * и предоставляет примеры запросов для каждого доступного источника.
 *
 * @returns {object} Result object with ok status and sources array
 *   {ok: true, sources: [{name, format, active, label}, ...]}
 *   OR {ok: false, error: {code: string, message: string}}
 *
 * @prerequisites
 * 1. Open SQLite Notebook tab in R7 Code plugin
 * 2. Load one or more files (CSV, XLSX, JSON, etc.)
 * 3. Run this macro
 *
 * @see https://r7-consult.ru/docs/api/sqlnotebook/getSources
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

        // ========== GET SOURCES / ПОЛУЧЕНИЕ ИСТОЧНИКОВ ==========

        console.log('[SqlNotebook] Retrieving loaded sources');
        var sourcesResult = Api.SqlNotebook.getSources();

        // ========== ERROR HANDLING / ОБРАБОТКА ОШИБОК ==========

        if (!sourcesResult.ok) {
            // Display error
            // Отобразить ошибку
            oWorksheet.GetRange("A1").SetValue("ERROR: " + sourcesResult.error.code);
            oWorksheet.GetRange("A2").SetValue(sourcesResult.error.message);

            var errorColor = Api.CreateColorFromRGB(255, 0, 0);
            oWorksheet.GetRange("A1").SetFontColor(errorColor);
            oWorksheet.GetRange("A1").SetBold(true);

            console.error('[SqlNotebook] getSources failed:', sourcesResult.error);
            return;
        }

        // ========== PROCESS SOURCES / ОБРАБОТКА ИСТОЧНИКОВ ==========

        var sources = sourcesResult.sources;
        console.log('[SqlNotebook] Found', sources.length, 'sources');

        var currentRow = 1;

        // Main Title
        // Главный заголовок
        oWorksheet.GetRange("A" + currentRow).SetValue("Loaded Data Sources in SQLite Notebook");
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        oWorksheet.GetRange("A" + currentRow).SetFontSize(16);
        var titleColor = Api.CreateColorFromRGB(0, 80, 160);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(titleColor);
        currentRow += 2;

        // Check if any sources are loaded
        // Проверка наличия загруженных источников
        if (sources.length === 0) {
            oWorksheet.GetRange("A" + currentRow).SetValue("⚠ No data sources loaded.");
            oWorksheet.GetRange("A" + (currentRow + 1)).SetValue("Please load files in SQLite Notebook first.");

            var warningColor = Api.CreateColorFromRGB(255, 140, 0);
            oWorksheet.GetRange("A" + currentRow).SetFontColor(warningColor);
            oWorksheet.GetRange("A" + currentRow).SetBold(true);

            console.warn('[SqlNotebook] No sources available');
            return;
        }

        // Summary
        // Итоги
        oWorksheet.GetRange("A" + currentRow).SetValue("Total sources: " + sources.length);
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        currentRow++;

        // Count active sources
        // Подсчет активных источников
        var activeCount = 0;
        sources.forEach(function(source) {
            if (source.active) activeCount++;
        });

        oWorksheet.GetRange("A" + currentRow).SetValue("Active sources: " + activeCount);
        var successColor = Api.CreateColorFromRGB(0, 150, 0);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(successColor);
        currentRow += 2;

        // ========== SOURCES TABLE / ТАБЛИЦА ИСТОЧНИКОВ ==========

        // Table headers
        // Заголовки таблицы
        var headers = ["Name", "Format", "Status", "Label"];
        for (var col = 0; col < headers.length; col++) {
            var cell = oWorksheet.GetRange(String.fromCharCode(65 + col) + currentRow);
            cell.SetValue(headers[col]);
            cell.SetBold(true);

            var headerBg = Api.CreateColorFromRGB(200, 220, 240);
            cell.SetFillColor(headerBg);
        }
        currentRow++;

        // Table rows - list all sources
        // Строки таблицы - список всех источников
        sources.forEach(function(source) {
            // Name (Column A)
            // Имя (Столбец A)
            oWorksheet.GetRange("A" + currentRow).SetValue(source.name);

            // Format (Column B)
            // Формат (Столбец B)
            oWorksheet.GetRange("B" + currentRow).SetValue(source.format.toUpperCase());

            // Status (Column C)
            // Статус (Столбец C)
            var status = source.active ? "✅ Active" : "⭕ Inactive";
            oWorksheet.GetRange("C" + currentRow).SetValue(status);

            if (source.active) {
                var activeColor = Api.CreateColorFromRGB(0, 150, 0);
                oWorksheet.GetRange("C" + currentRow).SetFontColor(activeColor);
            }

            // Label (Column D)
            // Метка (Столбец D)
            oWorksheet.GetRange("D" + currentRow).SetValue(source.label);

            currentRow++;
        });

        currentRow += 2;

        // ========== AUTOCOMPLETE TIP / ПОДСКАЗКА ПО АВТОДОПОЛНЕНИЮ ==========

        oWorksheet.GetRange("A" + currentRow).SetValue("💡 Autocomplete Tip:");
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        var tipColor = Api.CreateColorFromRGB(255, 140, 0);
        oWorksheet.GetRange("A" + currentRow).SetFontColor(tipColor);
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("When writing SQL queries in the macro editor:");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("1. Type: var result = Api.SqlNotebook.query('SELECT * FROM ");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("2. After FROM, press Ctrl+Space or just start typing");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("3. You'll see autocomplete suggestions for all loaded sources!");
        currentRow += 2;

        // ========== EXAMPLE QUERIES / ПРИМЕРЫ ЗАПРОСОВ ==========

        oWorksheet.GetRange("A" + currentRow).SetValue("Example Queries:");
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        oWorksheet.GetRange("A" + currentRow).SetFontSize(12);
        currentRow++;

        sources.forEach(function(source, index) {
            // Generate example query for each source
            // Создание примера запроса для каждого источника
            var exampleQuery = 'var result' + (index + 1) + ' = Api.SqlNotebook.query(\'SELECT * FROM "' + source.name + '" LIMIT 10\', "' + source.name + '");';

            oWorksheet.GetRange("A" + currentRow).SetValue(exampleQuery);
            oWorksheet.GetRange("A" + currentRow).SetFontSize(9);

            var codeColor = Api.CreateColorFromRGB(80, 80, 80);
            oWorksheet.GetRange("A" + currentRow).SetFontColor(codeColor);

            // Code background
            // Фон кода
            var codeBg = Api.CreateColorFromRGB(240, 240, 240);
            oWorksheet.GetRange("A" + currentRow).SetFillColor(codeBg);

            currentRow++;
        });

        currentRow += 2;

        // ========== USAGE EXAMPLE / ПРИМЕР ИСПОЛЬЗОВАНИЯ ==========

        oWorksheet.GetRange("A" + currentRow).SetValue("Dynamic Query Example:");
        oWorksheet.GetRange("A" + currentRow).SetBold(true);
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("// Automatically query all active sources:");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("sources.forEach(function(source) {");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("  if (source.active) {");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue('    var result = Api.SqlNotebook.query(\'SELECT COUNT(*) FROM "\' + source.name + \'"\', source.name);');
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("    // Process result...");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("  }");
        currentRow++;

        oWorksheet.GetRange("A" + currentRow).SetValue("});");
        currentRow += 2;

        // ========== FINAL MESSAGE / ФИНАЛЬНОЕ СООБЩЕНИЕ ==========

        oWorksheet.GetRange("A" + currentRow).SetValue("✅ Sources listed successfully!");
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
