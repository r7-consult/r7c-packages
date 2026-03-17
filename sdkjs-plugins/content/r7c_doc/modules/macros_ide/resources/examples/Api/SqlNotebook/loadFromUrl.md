/**
 * @file SqlNotebook_loadFromUrl_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.SqlNotebook.loadFromUrl
 * @author R7-Consult
 * @version 1.0.0
 * @date February 4, 2026
 *
 * @description
 * This macro demonstrates how to load a file from URL into SQLite Notebook
 * and immediately query it from the same macro.
 */

(function() {
    'use strict';

    try {
        var sheet = Api.GetActiveSheet();

        // Use a literal local path so pre-exec can resolve it (no variables).
        // Example: 'C:\\path\\to\\MOCK_DATA.csv' or 'file:///C:/path/to/MOCK_DATA.csv'
        var url = 'C:\\path\\to\\MOCK_DATA.csv';
        var loadResult = Api.SqlNotebook.loadFromUrl(url, {
            fileName: 'MOCK_DATA.csv',
            attach: true
        });

        if (!loadResult.ok) {
            sheet.GetRange('A1').SetValue('Load error: ' + loadResult.error.message);
            return;
        }

        var result = Api.SqlNotebook.query(
            'SELECT * FROM "MOCK_DATA.csv" LIMIT 5',
            'MOCK_DATA.csv'
        );

        if (!result.ok) {
            sheet.GetRange('A1').SetValue('Query error: ' + result.error.message);
            return;
        }

        // Output headers
        for (var i = 0; i < result.result.columns.length; i++) {
            sheet.GetRange(String.fromCharCode(65 + i) + '1').SetValue(result.result.columns[i]);
        }

        // Output rows
        for (var r = 0; r < result.result.rows.length; r++) {
            for (var c = 0; c < result.result.rows[r].length; c++) {
                sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(result.result.rows[r][c]);
            }
        }
    } catch (error) {
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            var ws = Api.GetActiveSheet();
            if (ws) {
                ws.GetRange('A1').SetValue('Error: ' + (error && error.message ? error.message : String(error)));
            }
        }
    }
})();
