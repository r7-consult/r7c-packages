/**
 * @file GetPivotByName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetPivotByName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a pivot table by its name.
 * It creates a pivot table, adds data, and then retrieves the pivot table by its name to add fields and data fields.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить сводную таблицу по ее имени.
 * Он создает сводную таблицу, добавляет данные, а затем извлекает сводную таблицу по ее имени, чтобы добавить поля и поля данных.
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
        // This example shows how to get a pivot table by its name.
        
        // How to find a pivot table.
        
        // Get a pivot table and by its name and update its fields.
        
        let worksheet = Api.GetActiveSheet();
        
        worksheet.GetRange('B1').SetValue('Region');
        worksheet.GetRange('C1').SetValue('Price');
        worksheet.GetRange('B2').SetValue('East');
        worksheet.GetRange('B3').SetValue('West');
        worksheet.GetRange('C2').SetValue(42.5);
        worksheet.GetRange('C3').SetValue(35.2);
        
        let dataRef = Api.GetRange("'Sheet1'!$B$1:$C$3");
        let pivotTable = Api.InsertPivotNewWorksheet(dataRef);
        
        Api.GetActiveSheet().GetPivotByName(pivotTable.GetName()).AddFields({
            rows: 'Region',
        });
        
        Api.GetActiveSheet().GetPivotByName(pivotTable.GetName()).AddDataField('Price');
        
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
