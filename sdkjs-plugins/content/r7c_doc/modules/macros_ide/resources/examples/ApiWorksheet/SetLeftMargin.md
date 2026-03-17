/**
 * @file SetLeftMargin_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetLeftMargin
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the left margin of the sheet.
 * It sets the left margin to 20.8 mm and then displays the updated margin.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить левое поле листа.
 * Он устанавливает левое поле на 20,8 мм, а затем отображает обновленное поле.
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
        // This example sets the left margin of the sheet.
        
        // How to set margin of the left side.
        
        // Resize the left margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetLeftMargin(20.8);
        let leftMargin = worksheet.GetLeftMargin();
        worksheet.GetRange("A1").SetValue("Left margin: " + leftMargin + " mm");
        
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
