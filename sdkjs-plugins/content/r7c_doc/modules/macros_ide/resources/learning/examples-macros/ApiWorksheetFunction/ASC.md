/**
 * @file ASC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.ASC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to change full-width (double-byte) characters to half-width (single-byte) characters for double-byte character set (DBCS) languages using ApiWorksheetFunction.ASC.
 * It converts "text" to its half-width equivalent and displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как изменить полноширинные (двухбайтовые) символы на полуширинные (однобайтовые) символы для языков с двухбайтовым набором символов (DBCS) с помощью ApiWorksheetFunction.ASC.
 * Он преобразует «текст» в его полуширинный эквивалент и отображает результат в ячейке A1.
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
        // This example shows for double-byte character set (DBCS) languages, the function changes full-width (double-byte) characters to half-width (single-byte) characters.
        
        // How to make characters half-width (single-byte) characters.
        
        // Use function to make characters half-width.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ASC("text"));
        
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
