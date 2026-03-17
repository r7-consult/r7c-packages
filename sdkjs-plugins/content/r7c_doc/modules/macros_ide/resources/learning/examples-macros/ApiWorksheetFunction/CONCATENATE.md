/**
 * @file CONCATENATE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CONCATENATE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to combine multiple text strings into one text string using ApiWorksheetFunction.CONCATENATE.
 * It concatenates "John", " ", and "Adams" and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как объединить несколько текстовых строк в одну текстовую строку с помощью ApiWorksheetFunction.CONCATENATE.
 * Он объединяет «John», « » и «Adams» и отображает результат в ячейке A1.
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
        // This example shows how to combine multiple text strings into one text string.
        
        // How to add multiple text strings into one text string.
        
        // Use function to create one text string from multiple ones.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CONCATENATE("John", " ", "Adams"));
        
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
