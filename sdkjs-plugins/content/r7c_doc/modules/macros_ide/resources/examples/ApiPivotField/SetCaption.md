/**
 * @file SetCaption_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotField.SetCaption
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a caption for a pivot field.
 * It creates a pivot table, adds fields, displays its initial caption, sets a new caption, and then displays the updated caption.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить заголовок для поля сводной таблицы.
 * Он создает сводную таблицу, добавляет поля, отображает его исходный заголовок, устанавливает новый заголовок, а затем отображает обновленный заголовок.
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
        // This example shows how to set a pivot field caption.
        
        // How to rename a pivot field caption.
        
        // Create a pivot table, add data to it then rename a specified pivot field's caption.
        
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
            rows: ['Region', 'Style'],
        });
        pivotTable.SetRowAxisLayout("Tabular", false);
        
        pivotTable.AddDataField('Price');
        
        let pivotWorksheet = Api.GetActiveSheet();
        let pivotField = pivotTable.GetPivotFields('Style');
        
        pivotWorksheet.GetRange('A12').SetValue('Style field caption');
        pivotWorksheet.GetRange('B12').SetValue(pivotField.GetCaption());
        
        pivotWorksheet.GetRange('A14').SetValue('New Style field caption');
        pivotField.SetCaption('My caption');
        pivotWorksheet.GetRange('B14').SetValue(pivotField.GetCaption());
        
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
