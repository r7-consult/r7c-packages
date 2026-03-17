/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.XOR
 * 
 *  Демонстрация использования метода XOR класса ApiWorksheetFunction
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
        // This example shows how to return the logical Exclusive Or value of all arguments. The function returns true when the number of true inputs is odd and false when the number of true inputs is even.
        
        // How to return the logical Exclusive Or value of all arguments.
        
        // Use a function to calculate Exclusive Or.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.TRUE(); //returns TRUE, doesnt require arguments
        
        worksheet.GetRange("A1").SetValue(ans);
        
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
