/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.RANK
 * 
 *  Демонстрация использования метода RANK класса ApiWorksheetFunction
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
        // This example shows how to return the rank of a number in a list of numbers: its size relative to other values in the list.
        
        // How to estimate the rank of a number in a list of numbers.
        
        // Use a function to estimate rank of the a number from the list.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [7,6,5,5];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        //method params
        let number = worksheet.GetRange("A3");
        let range = worksheet.GetRange("A1:A4");
        let order = 0;
        
        let func = Api.GetWorksheetFunction();
        let ans = func.RANK(number,range,order); 
        
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
