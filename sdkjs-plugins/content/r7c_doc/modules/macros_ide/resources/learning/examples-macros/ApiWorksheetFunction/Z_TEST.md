/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.Z_TEST
 * 
 *  Демонстрация использования метода Z_TEST класса ApiWorksheetFunction
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
        // This example shows how to return the one-tailed P-value of a z-test.
        
        // How to return one-tailed P-value.
        
        // Use a function to get one-tailed P-value.
        
        
        let worksheet = Api.GetActiveSheet();
        let argumentsArr = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16];
        
        // Place the numbers in cells
        for (let i = 0; i < argumentsArr.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(argumentsArr[i]);
        }
        
        // Get values from the range
        let data = worksheet.GetRange("A1:A16");
        
        // Calculate the TRIMMEAN of the range A1:A16
        let func = Api.GetWorksheetFunction();
        let result = func.Z_TEST(data, 4);
        worksheet.GetRange("B1").SetValue(result);
        
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
