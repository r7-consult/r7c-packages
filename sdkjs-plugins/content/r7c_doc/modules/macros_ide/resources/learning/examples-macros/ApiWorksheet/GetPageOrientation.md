/**
 * @file GetPageOrientation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetPageOrientation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the page orientation of the sheet.
 * It retrieves the page orientation and displays it in cell C1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить ориентацию страницы листа.
 * Он извлекает ориентацию страницы и отображает ее в ячейке C1.
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
        // This example shows how to get the page orientation.
        
        // How to get orientation of the sheet.
        
        // Get a sheet orientation.
        
        let worksheet = Api.GetActiveSheet();
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
