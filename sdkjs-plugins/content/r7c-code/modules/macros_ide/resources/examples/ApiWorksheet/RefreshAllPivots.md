/**
 * @file RefreshAllPivots_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.RefreshAllPivots
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to refresh all pivot tables on the sheet.
 * It creates multiple pivot tables, adds data fields to them, and then refreshes all pivot tables.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как обновить все сводные таблицы на листе.
 * Он создает несколько сводных таблиц, добавляет в них поля данных, а затем обновляет все сводные таблицы.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example shows how to refresh all pivot tables.
        
        // How to refresh field values of all pivot tables.
        
        // Refresh pivot tables from the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        let pivotRef = worksheet.GetRange('A7');
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('A7'));
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('D7'));
        Api.InsertPivotExistingWorksheet(dataRef, worksheet.GetRange('G7'));
        
        worksheet.GetAllPivotTables().forEach(function (pivot) {
            pivot.AddDataField('Price');
        });
        
        worksheet.RefreshAllPivots();
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
