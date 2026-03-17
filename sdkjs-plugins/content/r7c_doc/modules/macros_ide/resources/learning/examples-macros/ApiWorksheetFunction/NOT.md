/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NOT
 * 
 *  Демонстрация использования метода NOT класса ApiWorksheetFunction
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
        // This example shows how to сheck if the specified logical value is true or false. The function returns true if the argument is false and false if the argument is true.
        
        // How to negate a boolean value.
        
        // Use a function to get the opposite of the boolean value.
        
        const worksheet = Api.GetActiveSheet();
        
        let condition = 12 < 100;
        let func = Api.GetWorksheetFunction();
        let ans = func.NOT(condition);
        
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
