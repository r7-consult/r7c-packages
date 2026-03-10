/**
 * @file GetTableRange1_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiPivotTable.GetTableRange1
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the range of the entire pivot table report, excluding page fields.
 * It creates a pivot table, adds data and fields, and then selects the table range without page fields.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить диапазон всего отчета сводной таблицы, исключая поля страниц.
 * Он создает сводную таблицу, добавляет данные и поля, а затем выбирает диапазон таблицы без полей страниц.
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
        // Returns a Range object that represents the entire pivot table report, but doesn't include page fields.
        
        // How to get a pivot table range without page fields.
        
        // Create the ApiPivotTable object, add data to it then get the ApiRange object for the entire pivot table report but without page fields.
        
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
            pages: 'Style',
        });
        
        pivotTable.AddDataField('Price');
        
        pivotTable.GetTableRange1().Select();
        
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
