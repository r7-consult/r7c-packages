/**
 * @file CODE_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.CODE
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the code number from your computer's character set for the first character in the specified text string.
 * It gets the code number for the first character of "office" and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть кодовый номер из набора символов вашего компьютера для первого символа в указанной текстовой строке.
 * Он получает кодовый номер для первого символа «office» и отображает результат в ячейке A1.
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
        // This example shows how to return the code number from your computer's character set for the first character in the specified text string.
        
        // How to return the code number from your computer's character set.
        
        // Use function to get a code number from your computer's character set.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.CODE("office"));
        
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
