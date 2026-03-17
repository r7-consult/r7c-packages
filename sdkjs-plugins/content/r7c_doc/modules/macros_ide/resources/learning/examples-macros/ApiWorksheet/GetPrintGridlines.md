/**
 * @file GetPrintGridlines_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetPrintGridlines
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the page PrintGridlines property which specifies whether the sheet gridlines must be printed or not.
 * It sets the print gridlines to true and then displays the setting.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство PrintGridlines страницы, которое указывает, должны ли печататься линии сетки листа.
 * Он устанавливает печать линий сетки в значение true, а затем отображает настройку.
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
        // This example shows how to get the page PrintGridlines property which specifies whether the sheet gridlines must be printed or not.
        
        // How to find out whether sheet gridlines should be printed or not.
        
        // Get a boolean value representing whether to print gridlines or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPrintGridlines(true);
        worksheet.GetRange("A1").SetValue("Gridlines of cells will be printed on this page: " + worksheet.GetPrintGridlines());
        
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
