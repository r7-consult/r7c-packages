/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.WEEKNUM
 * 
 *  Демонстрация использования метода WEEKNUM класса ApiWorksheetFunction
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
        // This example shows how to return a number from 1 to 7 identifying the day of the week of the specified date.
        
        // How to return a weekday.
        
        // Use a function to get a weekday using numbers.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.WEEKNUM("11/5/2018", 2); 
        
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
