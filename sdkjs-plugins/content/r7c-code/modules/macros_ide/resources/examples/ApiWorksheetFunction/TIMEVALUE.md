/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TIMEVALUE
 * 
 *  Демонстрация использования метода TIMEVALUE класса ApiWorksheetFunction
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
        // This example shows how to convert a text time to a serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula.
        
        // How to create a serial number from a date time object.
        
        // Use a function to convert date, hours, minutes and seconds to serial numbers.
        
        const worksheet = Api.GetActiveSheet();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.TIMEVALUE("11/5/18 11:17:00 am"); 
        
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
