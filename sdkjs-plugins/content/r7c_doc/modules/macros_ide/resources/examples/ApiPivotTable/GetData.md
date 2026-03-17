/**
 * @file GetData_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotTable.GetData
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get data from a pivot table using names.
 * It creates a pivot table, adds data and fields, and then retrieves and displays a specific data point.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить данные из сводной таблицы по именам.
 * Он создает сводную таблицу, добавляет данные и поля, а затем извлекает и отображает определенную точку данных.
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
        // This example shows how to get data from a pivot table using names.
        
        // How to get a field data from a pivot table.
        
        // Create a pivot table, add data to it then get some data from it.
        
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
        let dataField = pivotTable.AddDataField('Price');
        dataField.SetName('Regional prices');
        pivotTable.AddFields({
            rows: 'Region',
            columns: 'Style',
        });
        
        let pivotWorksheet = Api.GetActiveSheet();
        
        let value = pivotTable.GetData(['East', 'Tee']);
        pivotWorksheet.GetRange('A9').SetValue('Value of East Tee Price:');
        pivotWorksheet.GetRange('B9').SetValue(value);
        
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
