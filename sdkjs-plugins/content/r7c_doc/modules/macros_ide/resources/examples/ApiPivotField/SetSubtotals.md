/**
 * @file SetSubtotals_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotField.SetSubtotals
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the subtotals for a pivot field.
 * It creates a pivot table, adds fields, sets the subtotals for the 'Region' pivot field to include 'Count', and then displays the subtotals.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить промежуточные итоги для поля сводной таблицы.
 * Он создает сводную таблицу, добавляет поля, устанавливает промежуточные итоги для поля сводной таблицы «Регион» для включения «Count», а затем отображает промежуточные итоги.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example shows how to set subtotals of a pivot field.
        
        // How to change a pivot field subtotals.
        
        // Create a pivot table, add data to it then set subtotals of a specified pivot.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Style');
        worksheet.GetRange('D1').SetValue('Price');
        
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('B4').SetValue('East');
        worksheet.GetRange('B5').SetValue('West');
        
        worksheet.GetRange('C2').SetValue('Fancy');
        worksheet.GetRange('C3').SetValue('Fancy');
        worksheet.GetRange('C4').SetValue('Tee');
        worksheet.GetRange('C5').SetValue('Tee');
        
        worksheet.GetRange('D2').SetValue(42.5);
        worksheet.GetRange('D3').SetValue(35.2);
        worksheet.GetRange('D4').SetValue(12.3);
        worksheet.GetRange('D5').SetValue(24.8);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$D$5");
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
        pivotTable.AddFields({
            columns: ['Style'],
            rows: 'Region',
        });
        
        pivotTable.AddDataField('Price');
        
        let pivotWorksheet = Api.GetActiveSheet();
        let pivotField = pivotTable.GetPivotFields('Region');
        
        pivotField.SetSubtotals({
            Count: true,
        });
        
        let subtotals = pivotField.GetSubtotals();
        pivotWorksheet.GetRange('A11').SetValue('Region subtotals');
        let k = 12;
        for (let i in subtotals) {
            pivotWorksheet.GetRangeByNumber(k, 0).SetValue(i);
            pivotWorksheet.GetRangeByNumber(k++, 1).SetValue(subtotals[i]);
        }
        
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
