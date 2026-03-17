/**
 * @file SetHyperlink_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.SetHyperlink
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a hyperlink to the specified range.
 * It adds a hyperlink to cell A1 with a URL, display text, and screen tip.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить гиперссылку к указанному диапазону.
 * Он добавляет гиперссылку к ячейке A1 с URL-адресом, отображаемым текстом и всплывающей подсказкой.
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
        // This example adds a hyperlink to the specified range.
        
        // How to add hyperlinks to the range.
        
        // Add a hyperlink to the cell.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.SetHyperlink("A1", "https://api.R7 Office.com/docbuilder/basic", "Api R7 Office", "R7 Office for developers");
        
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
