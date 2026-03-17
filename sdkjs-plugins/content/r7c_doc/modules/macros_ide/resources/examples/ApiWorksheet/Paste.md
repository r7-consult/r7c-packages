/**
 * @file Paste_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.Paste
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to paste data from the clipboard.
 * It pastes the content from the clipboard to the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить данные из буфера обмена.
 * Он вставляет содержимое из буфера обмена в активный лист.
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
        // This example shows how to paste date from clipboard.
        
        // How to paste a copied or cut data from the clipboard.
        
        // Paste to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.Paste();
        
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
