/**
 * @file SetName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set a name to the active sheet.
 * It renames the active worksheet to "sheet 1" and then displays the new name.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить имя активного листа.
 * Он переименовывает активный лист в «лист 1», а затем отображает новое имя.
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
        // This example sets a name to the active sheet.
        
        // How to set name of the sheet.
        
        // Rename the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetName("sheet 1");
        let name = worksheet.GetName();
        worksheet.GetRange("A1").SetValue("Worksheet name: ");
        worksheet.GetRange("A1").AutoFit(false, true);
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
