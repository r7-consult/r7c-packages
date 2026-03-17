/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TRUE
 * 
 *  Демонстрация использования метода TRUE класса ApiWorksheetFunction
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
        // This example shows how to return the true logical value.
        
        // How to get a true value.
        
        // Use a function to return true value.
        
        const worksheet = Api.GetActiveSheet();
        
        let logical1 = 1 > 0;
        let logical2 = 2 > 0;
        
        let func = Api.GetWorksheetFunction();
        let ans = func.XOR(logical1, logical2); //Works on XOR gate logic
        
        worksheet.GetRange("C1").SetValue(ans);
        
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
