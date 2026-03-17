/**
 * @file GetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a sheet name.
 * It retrieves the name of the active worksheet and displays it in cell B1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить имя листа.
 * Он извлекает имя активного листа и отображает его в ячейке B1.
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
        // This example shows how to get a sheet name.
        
        // How to get name of the sheet.
        
        // Get a sheet name.
        
        let worksheet = Api.GetActiveSheet();
        let name = worksheet.GetName();
        worksheet.GetRange("A1").SetValue("Name: ");
        worksheet.GetRange("B1").SetValue(name);
        
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
