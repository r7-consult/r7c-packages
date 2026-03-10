/**
 * @file SqlNotebook_Paging_Example.js
 * @brief Demonstrates paging options in Api.SqlNotebook.query
 * @date February 4, 2026
 */

(function() {
    'use strict';

    var sheet = Api.GetActiveSheet();
    var sql = 'SELECT * FROM "MOCK_DATA.csv"';

    var page0 = Api.SqlNotebook.query(sql, 'MOCK_DATA.csv', {
        paging: { page: 0, pageSize: 5 }
    });

    var page1 = Api.SqlNotebook.query(sql, 'MOCK_DATA.csv', {
        paging: { page: 1, pageSize: 5 }
    });

    if (!page0.ok) {
        sheet.GetRange('A1').SetValue('Page 1 error: ' + page0.error.message);
        return;
    }

    if (!page1.ok) {
        sheet.GetRange('A1').SetValue('Page 2 error: ' + page1.error.message);
        return;
    }

    // Headers
    for (var i = 0; i < page0.result.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + '1').SetValue(page0.result.columns[i]);
    }

    // Page 1 rows
    for (var r = 0; r < page0.result.rows.length; r++) {
        for (var c = 0; c < page0.result.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(page0.result.rows[r][c]);
        }
    }

    var offsetRow = page0.result.rows.length + 4;
    sheet.GetRange('A' + offsetRow).SetValue('Page 2');

    // Page 2 rows
    for (var r2 = 0; r2 < page1.result.rows.length; r2++) {
        for (var c2 = 0; c2 < page1.result.rows[r2].length; c2++) {
            sheet.GetRange(String.fromCharCode(65 + c2) + (offsetRow + r2 + 1)).SetValue(page1.result.rows[r2][c2]);
        }
    }
})();