/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of an ApiRange object.
 * It sets a value in cell A1, gets the range, and then displays its class type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiRange.
 * Он устанавливает значение в ячейке A1, получает диапазон, а затем отображает его тип класса.
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
        // This example gets a class type and inserts it into the table.
        
        // How to get a class type of ApiRange.
        
        // Get a class type of ApiRange and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRange("A1");
        range.SetValue("This is just a sample text in the cell A1.");
        let classType = range.GetClassType();
        worksheet.GetRange('A3').SetValue("Class type: " + classType);
        
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
