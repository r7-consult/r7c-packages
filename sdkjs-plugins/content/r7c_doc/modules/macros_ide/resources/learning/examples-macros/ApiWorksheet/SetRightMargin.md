/**
 * @file SetRightMargin_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetRightMargin
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the right margin of the sheet.
 * It sets the right margin to 20.8 mm and then displays the updated margin.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить правое поле листа.
 * Он устанавливает правое поле на 20,8 мм, а затем отображает обновленное поле.
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
        // This example sets the right margin of the sheet.
        
        // How to set margin of the right side.
        
        // Resize the right margin of the sheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetRightMargin(20.8);
        let rightMargin = worksheet.GetRightMargin();
        worksheet.GetRange("A1").SetValue("Right margin: " + rightMargin + " mm");
        
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
