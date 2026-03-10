/**
 * @file Move_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotField.Move
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to move a pivot field.
 * It creates a pivot table, adds fields, and then moves the 'Region' pivot field to the columns area after a delay.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как переместить поле сводной таблицы.
 * Он создает сводную таблицу, добавляет поля, а затем перемещает поле сводной таблицы «Регион» в область столбцов с задержкой.
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
        // This example shows how to move a pivot field.
        
        // How to change the position a pivot field.
        
        // Create a pivot table, add data to it then move a specified pivot field by columns.
        
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
            rows: 'Region',
            columns: 'Style',
        });
        
        let pivotWorksheet = Api.GetActiveSheet();
        pivotTable.AddDataField('Price');
        let pivotField = pivotTable.GetPivotFields('Region');
        pivotWorksheet.GetRange('A10').SetValue('The Region field will be moved soon');
        
        setTimeout(function () {
            pivotField.Move('Columns');
        }, 5000);
        
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
