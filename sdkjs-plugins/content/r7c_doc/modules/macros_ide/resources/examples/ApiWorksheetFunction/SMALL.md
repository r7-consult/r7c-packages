/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.SMALL
 * 
 *  Демонстрация использования метода SMALL класса ApiWorksheetFunction
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
        // This example shows how to return the k-th smallest value in a data set. For example, the fifth smallest number.
        
        // How to return the k-th smallest value from data set.
        
        // Use a function to get the smallest value from data set indicated.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [1, 0, 0, 0, 0, 1, 0, 0, 2, 3, 4, 5, 6, 8, 10, 12];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        // method params
        let range = worksheet.GetRange("A1:A16");
        let position = 8;
        
        let func = Api.GetWorksheetFunction();
        let ans = func.SMALL(range, position);
        
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
