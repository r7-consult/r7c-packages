/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.MEDIAN
 * 
 *  Демонстрация использования метода MEDIAN класса ApiWorksheetFunction
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
        // This example shows how to return the median, or the number in the middle of the set of given numbers.
        
        // How to get a median from the list.
        
        // Use a function to get a value that located in the middle of the list.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let median = func.MEDIAN(4,45,12,34,3,54,2,2);
        worksheet.GetRange("C1").SetValue(median);
        
        
        
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
