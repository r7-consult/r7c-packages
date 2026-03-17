/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISNONTEXT
 * 
 *  Демонстрация использования метода ISNONTEXT класса ApiWorksheetFunction
 * https://r7-consult.ru/
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
        // This example shows how to check whether a value is not text (blank cells are not text), and returns true or false. 
        
        // How to check if the cell contains a non-text value.
        
        // Use a function to check whether a range data is a text or not.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ISNONTEXT("#N/A"));
        worksheet.GetRange("A2").SetValue(func.ISNONTEXT(255));
        worksheet.GetRange("A3").SetValue(func.ISNONTEXT("Online Office"));
        
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
