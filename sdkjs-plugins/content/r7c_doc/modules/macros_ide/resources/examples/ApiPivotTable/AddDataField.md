/**
 * @file AddDataField_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotTable.AddDataField
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a data field to a pivot table.
 * It creates a pivot table, moves the 'Region' field to rows, and then adds a 'Price' data field and renames it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить поле данных в сводную таблицу.
 * Он создает сводную таблицу, перемещает поле «Регион» в строки, а затем добавляет поле данных «Цена» и переименовывает его.
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
        // This example shows how to a data field to a pivot table.
        
        // How to add new field to the table.
        
        // Create a pivot table, add data to it then add new data field to it.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
        pivotTable.MoveField('Region', 'Rows');
        
        let dataField = pivotTable.AddDataField('Price');
        dataField.SetName('Regional prices');
        
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
