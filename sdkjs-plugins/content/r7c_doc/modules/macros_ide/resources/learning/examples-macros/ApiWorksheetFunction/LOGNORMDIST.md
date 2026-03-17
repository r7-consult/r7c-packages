/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.LOGNORMDIST
 * 
 *  Демонстрация использования метода LOGNORMDIST класса ApiWorksheetFunction
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
        // This example shows how to return the cumulative lognormal distribution of x, where ln(x) is normally distributed with the specified parameters.
        
        // How to get the cumulative lognormal distribution of x.
        
        // Use a function to return the cumulative lognormal distribution.
        
        const worksheet = Api.GetActiveSheet();
        
        //configure function parameters
        let numbersArr = [4, 7, 40];
        
        //set values in cells
        for (let i = 0; i < numbersArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(numbersArr[i]);
        }
        
        //get parameters
        let xValue = worksheet.GetRange("A1");
        let mean = worksheet.GetRange("A2");
        let standardDeviation = worksheet.GetRange("A3");
        
        //invoke LOGNORMDIST method
        let func = Api.GetWorksheetFunction();
        let ans = func.LOGNORMDIST(xValue, mean, standardDeviation);
        
        //print answer
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
