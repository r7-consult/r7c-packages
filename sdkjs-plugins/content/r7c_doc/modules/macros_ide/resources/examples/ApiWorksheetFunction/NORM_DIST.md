/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NORM_DIST
 * 
 *  Демонстрация использования метода NORM_DIST класса ApiWorksheetFunction
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
        // This example shows how to return the normal distribution for the specified mean and standard deviation.
        
        // How to calculate the normal distribution.
        
        // Use a function to get the normal distribution knowing the mean and standard deviation.
        
        const worksheet = Api.GetActiveSheet();
        let valueArr = [36, 6, 7, false];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        //method params
        let x = worksheet.GetRange("A1").GetValue();
        let mean = worksheet.GetRange("A2").GetValue();
        let standardDeviation = worksheet.GetRange("A3").GetValue();
        let cumulative = worksheet.GetRange("A4").GetValue();
        let func = Api.GetWorksheetFunction();
        let normalDist = func.NORM_DIST(x, mean, standardDeviation, cumulative);
        worksheet.GetRange("C1").SetValue(normalDist);
        
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
