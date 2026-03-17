/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NEGBINOMDIST
 * 
 *  Демонстрация использования метода NEGBINOMDIST класса ApiWorksheetFunction
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
        // This example shows how to return the negative binomial distribution, the probability that there will be the specified number of failures before the last success, with the specified probability of a success.
        
        // How to return the negative binomial distribution.
        
        // Use a function to get the probability of the specified number of failures before the last success (negative binomial distribution).
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [6, 32, 0.7];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        //method params
        let numberF = worksheet.GetRange("A1").GetValue();
        let numberS = worksheet.GetRange("A2").GetValue();
        let probabilityS = worksheet.GetRange("A3").GetValue();
        
        let func = Api.GetWorksheetFunction();
        let ans = func.NEGBINOMDIST(numberF, numberS, probabilityS);
        
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
