/**
 * @file GetPrintHeadings_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetPrintHeadings
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the page PrintHeadings property which specifies whether the sheet row/column headings must be printed or not.
 * It sets the print headings to true and then displays the setting.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство PrintHeadings страницы, которое указывает, должны ли печататься заголовки строк/столбцов листа.
 * Он устанавливает печать заголовков в значение true, а затем отображает настройку.
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
        // This example shows how to get the page PrintHeadings property which specifies whether the sheet row/column headings must be printed or not.
        
        // How to find out whether sheet headings should be printed or not.
        
        // Get a boolean value representing whether to print row and column headings or not.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetPrintHeadings(true);
        worksheet.GetRange("A1").SetValue("Row and column headings will be printed with this page: " + worksheet.GetPrintHeadings());
        
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
