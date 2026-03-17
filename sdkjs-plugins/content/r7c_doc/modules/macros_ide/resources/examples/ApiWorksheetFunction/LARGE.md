/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.LARGE
 * 
 *  Демонстрация использования метода LARGE класса ApiWorksheetFunction
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
        // This example shows how to return the k-th largest value in a data set. For example, the fifth largest number.
        
        // How to find the k-th largest value in a data set.
        
        // Use a function to find out the largest value in a data set specifying its order number.
        
        const worksheet = Api.GetActiveSheet();
        
        let numbersArr = [4, 13, 27, 56, 46, 79, 22, 12];
        
        // Place the numbers in cells
        
        for (let i = 0; i < numbersArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(numbersArr[i]);
        }
        
        let func = Api.GetWorksheetFunction();
        let range = worksheet.GetRange("A1:A8");
        let largePostion = 4;
        let kLargest = func.LARGE(range, largePostion);
        worksheet.GetRange("C1").SetValue(kLargest);
        
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
