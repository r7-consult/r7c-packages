/**
 * @file SetTopMargin_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetTopMargin
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the top margin of the sheet.
 * It sets the top margin to 25.1 mm and then displays the updated margin.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить верхнее поле листа.
 * Он устанавливает верхнее поле на 25,1 мм, а затем отображает обновленное поле.
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
        // This example sets the top margin of the sheet.
        
        // How to set margin of the top side.
        
        // Resize the top margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetTopMargin(25.1);
        let topMargin = worksheet.GetTopMargin();
        worksheet.GetRange("A1").SetValue("Top margin: " + topMargin + " mm");
        
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
