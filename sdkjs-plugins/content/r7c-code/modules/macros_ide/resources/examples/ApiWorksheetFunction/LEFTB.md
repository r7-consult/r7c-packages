/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.LEFTB
 * 
 *  Демонстрация использования метода LEFTB класса ApiWorksheetFunction
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
        // This example shows how to extract the substring from the specified string starting from the left character and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
        
        // How to extract the substring from the left character.
        
        // Use a function to get the substring from the left.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.LEFTB("Online Office", 6));
        
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
