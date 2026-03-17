/**
 * @file GetOrientation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetOrientation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the orientation of a range.
 * It sets values in a range, sets its orientation to "xlUpward", and then displays the orientation.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить ориентацию диапазона.
 * Он устанавливает значения в диапазоне, устанавливает его ориентацию на «xlUpward», а затем отображает ориентацию.
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
        // This example shows how to get the range angle.
        
        // How to find out cell orientation of a range.
        
        // Get a range, get its orientation (upward, downward, etc.) and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        let range = worksheet.GetRange("A1:B1");
        range.SetOrientation("xlUpward");
        let orientation = range.GetOrientation();
        worksheet.GetRange("A3").SetValue("Orientation: " + orientation);
        
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
