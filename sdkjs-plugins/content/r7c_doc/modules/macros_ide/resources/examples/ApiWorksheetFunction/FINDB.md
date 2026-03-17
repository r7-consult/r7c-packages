/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FINDB
 * 
 *  Демонстрация использования метода FINDB класса ApiWorksheetFunction
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
        // This example shows how to find the specified substring within another string and is intended for languages that use the double-byte character set (DBCS) like Japanese, Chinese, Korean etc.
        
        // How to get an index of a substring from a string and its double-byte character language.
        
        // Use function to find out a starting position of a substring containing in another text and its character set (DBCS).
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.FINDB("line", "Online Office"));
        
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
