/**
 * @file GetSubtotalName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotField.GetSubtotalName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the subtotal name of a pivot field.
 * It creates a pivot table, adds fields, sets a custom subtotal name for the 'Region' pivot field, and then displays it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить имя промежуточного итога поля сводной таблицы.
 * Он создает сводную таблицу, добавляет поля, устанавливает настраиваемое имя промежуточного итога для поля сводной таблицы «Регион», а затем отображает его.
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
        // This example shows how to get a subtotal name of a pivot field.
        
        // How to get a pivot field's subtotal name.
        
        // Create a pivot table, add data to it then get subtotal name of a specified pivot field.
        
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
            columns: ['Region', 'Style'],
        });
        
        pivotTable.AddDataField('Price');
        
        let pivotWorksheet = Api.GetActiveSheet();
        let pivotField = pivotTable.GetPivotFields('Region');
        
        pivotField.SetSubtotalName('My name');
        
        pivotWorksheet.GetRange('A14').SetValue('Region subtotal name');
        pivotWorksheet.GetRange('B14').SetValue(pivotField.GetSubtotalName());
        
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
