/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.WORKDAY
 * 
 *  Демонстрация использования метода WORKDAY класса ApiWorksheetFunction
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
        // This example shows how to return the serial number of the date before or after a specified number of workdays.
        
        // How to return the serial number of the date adding some workdays.
        
        // Use a function to calculate the serial number of the date.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.WORKDAY("12/7/1981", 56); 
        
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
