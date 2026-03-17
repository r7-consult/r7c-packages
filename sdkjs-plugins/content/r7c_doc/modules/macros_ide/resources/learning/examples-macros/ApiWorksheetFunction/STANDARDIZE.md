/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.STANDARDIZE
 * 
 *  Демонстрация использования метода STANDARDIZE класса ApiWorksheetFunction
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
        // This example shows how to return a normalised value from a distribution characterised by a mean and standard deviation.
        
        // How to calculate the normalised value from a distribution.
        
        // Use a function to get the normalised value from a distribution by different parameters.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [5, -2, 10];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        // method params
        let x = worksheet.GetRange("A1");
        let mean = worksheet.GetRange("A2");
        let stdDev = worksheet.GetRange("A3");
        
        let func = Api.GetWorksheetFunction();
        let ans = func.STANDARDIZE(x, mean, stdDev);
        
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
