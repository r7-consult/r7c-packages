/**
 * @file SetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotDataField.SetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a name for a pivot data field.
 * It creates a pivot table, adds a data field, displays its initial name, sets a new name, and then displays the updated name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить имя для поля данных сводной таблицы.
 * Он создает сводную таблицу, добавляет поле данных, отображает его исходное имя, устанавливает новое имя, а затем отображает обновленное имя.
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
        // This example shows how to set a name for data field.
        
        // How to rename a table element.
        
        // Create a pivot table, add data to it then set a custom data field's name.
        
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
        
        pivotTable.AddDataField('Price');
        
        let pivotWorksheet = Api.GetActiveSheet();
        let dataField = pivotTable.GetDataFields('Sum of Price');
        
        pivotWorksheet.GetRange('A12').SetValue('Data field name');
        pivotWorksheet.GetRange('B12').SetValue(dataField.GetName());
        
        dataField.SetName('My Sum of Price');
        pivotWorksheet.GetRange('A13').SetValue('New Data field name');
        pivotWorksheet.GetRange('B13').SetValue(dataField.GetName());
        
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
