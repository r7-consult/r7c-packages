/**
 * @file AddDefName_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.AddDefName
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a new name to the worksheet.
 * It sets values in a range, adds a defined name to it, and then displays a confirmation message.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить новое имя на лист.
 * Он устанавливает значения в диапазоне, добавляет к нему определенное имя, а затем отображает сообщение с подтверждением.
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
        // This example adds a new name to the worksheet.
        
        // How to change a name of the worksheet range.
        
        // Name a range from a worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.AddDefName("numbers", "Sheet1!$A$1:$B$1");
        worksheet.GetRange("A3").SetValue("We defined a name 'numbers' for a range of cells A1:B1.");
        
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
