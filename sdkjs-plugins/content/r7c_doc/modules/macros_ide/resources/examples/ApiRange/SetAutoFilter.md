/**
 * @file SetAutoFilter_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.SetAutoFilter
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set an autofilter by cell range.
 * It sets values in a range and then applies an autofilter to display only specific values.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить автофильтр по диапазону ячеек.
 * Он устанавливает значения в диапазоне, а затем применяет автофильтр для отображения только определенных значений.
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
        // This example sets the autofilter by cell range.
        
        // How to automatically filter the specified range values.
        
        // Automatically filter out a range values.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("header");
        worksheet.GetRange("A2").SetValue("value2");
        worksheet.GetRange("A3").SetValue("value3");
        worksheet.GetRange("A4").SetValue("value4");
        worksheet.GetRange("A5").SetValue("value5");
        let range = worksheet.GetRange("A1:A5");
        range.SetAutoFilter(1, ["value2","value3"], "xlFilterValues");
        
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
