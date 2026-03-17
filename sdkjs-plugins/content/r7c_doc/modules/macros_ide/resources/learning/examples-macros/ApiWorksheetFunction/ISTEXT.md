/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISTEXT
 * 
 *  Демонстрация использования метода ISTEXT класса ApiWorksheetFunction
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
        // This example shows how to check whether a value is text, and returns true or false.
        
        // How to know whether a value is a text.
        
        // Use a function to find out whether a value is a text.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.ISTEXT(255));
        worksheet.GetRange("A2").SetValue(func.ISTEXT("#N/A"));
        worksheet.GetRange("A3").SetValue(func.ISTEXT("Online Office"));
        
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
