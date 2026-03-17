/**
 * @file SetVisible_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetVisible
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the state of sheet visibility.
 * It sets the active worksheet to visible and then displays a message.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить состояние видимости листа.
 * Он устанавливает активный лист видимым, а затем отображает сообщение.
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
        // This example sets the state of sheet visibility.
        
        // How to set visibility of the sheet.
        
        // Make a sheet visible or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetVisible(true);
        worksheet.GetRange("A1").SetValue("The current worksheet is visible.");
        
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
