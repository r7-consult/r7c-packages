/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.YEARFRAC
 * 
 *  Демонстрация использования метода YEARFRAC класса ApiWorksheetFunction
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
        // This example shows how to return the year fraction representing the number of whole days between the start date and end date.
        
        // How to return the year fraction.
        
        // Use a function to calculate a year fraction.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.YEARFRAC("12/7/1981", "11/5/2018");
        
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
