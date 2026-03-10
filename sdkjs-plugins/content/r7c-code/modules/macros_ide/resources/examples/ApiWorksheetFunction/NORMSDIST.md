/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NORMSDIST
 * 
 *  Демонстрация использования метода NORMSDIST класса ApiWorksheetFunction
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
        // This example shows how to return the standard normal cumulative distribution (has a mean of zero and a standard deviation of one).
        
        // How to calculate the standard normal cumulative distribution.
        
        // Use a function to get the standard normal cumulative distribution.
        
        const worksheet = Api.GetActiveSheet();
        let valueArr = [0.34, 7, 3];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        //method params
        let probability = worksheet.GetRange("A1").GetValue();
        let mean = worksheet.GetRange("A2").GetValue();
        let standardDeviation = worksheet.GetRange("A3").GetValue();
        
        let func = Api.GetWorksheetFunction();
        let inv = func.NORM_INV(probability, mean, standardDeviation);
        worksheet.GetRange("C1").SetValue(inv);
        
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
