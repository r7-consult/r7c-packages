/**
 * @file GetAllPivotTables_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllPivotTables
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all pivot tables from the sheet.
 * It creates multiple pivot tables, and then iterates through them to add a data field to each.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все сводные таблицы с листа.
 * Он создает несколько сводных таблиц, а затем перебирает их, чтобы добавить поле данных к каждой.
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
        // This example shows how to get all pivot tables from the sheet.
        
        // How to get all pivot tables.
        
        // Get all pivot tables as an array.
        
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
