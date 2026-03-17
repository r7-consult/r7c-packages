/**
 * @file PasteSpecial_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.PasteSpecial
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to paste data from the clipboard with specified options.
 * It sets a value in cell A1, copies it, and then pastes it to cell B1 using PasteSpecial with specific options.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вставить данные из буфера обмена с указанными параметрами.
 * Он устанавливает значение в ячейке A1, копирует его, а затем вставляет в ячейку B1 с помощью PasteSpecial с определенными параметрами.
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
        // This example pastes data from theclipboard (if it is possible) with the specified options.
        
        // How to paste the copied or cut data from the clipboard using the special paste options.
        
        // Create a range, copy its value and paste it into another one with the specified properties.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("test");
        worksheet.GetRange("A1").Copy();
        worksheet.GetRange("B1").PasteSpecial("xlPasteAll", "xlPasteSpecialOperationNone", false, false);
        
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
