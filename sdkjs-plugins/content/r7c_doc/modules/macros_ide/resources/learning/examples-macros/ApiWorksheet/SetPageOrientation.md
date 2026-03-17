/**
 * @file SetPageOrientation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetPageOrientation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the page orientation.
 * It sets the page orientation to "xlPortrait" and then displays the updated orientation.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить ориентацию страницы.
 * Он устанавливает ориентацию страницы на «xlPortrait», а затем отображает обновленную ориентацию.
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
        // This example sets the page orientation.
        
        // How to change a page orientation.
        
        // Set a page orientation and display it in the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPageOrientation("xlPortrait");
        let pageOrientation = worksheet.GetPageOrientation();
        worksheet.GetRange("A1").SetValue("Page orientation: ");
        worksheet.GetRange("C1").SetValue(pageOrientation);
        
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
