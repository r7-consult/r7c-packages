/**
 * @file GetAreas_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetAreas
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a collection of ranges.
 * It sets a value in range B1:D1, selects it, gets its areas, and then displays the count of ranges in the areas.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить коллекцию диапазонов.
 * Он устанавливает значение в диапазоне B1:D1, выбирает его, получает его области, а затем отображает количество диапазонов в областях.
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
        // This example shows how to get a collection of the ranges.
        
        // How to get range areas.
        
        // Get range areas, count them and display the result in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        let areas = range.GetAreas();
        let count = areas.GetCount();
        range = worksheet.GetRange("A5");
        range.SetValue("The number of ranges in the areas: ");
        range.AutoFit(false, true);
        worksheet.GetRange("B5").SetValue(count);
        
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
