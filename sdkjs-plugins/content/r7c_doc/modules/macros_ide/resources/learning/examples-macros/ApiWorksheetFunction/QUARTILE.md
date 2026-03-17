/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.QUARTILE
 * 
 *  Демонстрация использования метода QUARTILE класса ApiWorksheetFunction
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
        // This example shows how to return the quartile of a data set.
        
        // How to get the quartile of a data set.
        
        // Use a function to calculate a part divided by four of a data set.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        let quart = 3; //eg: third quartile (75th percentile)
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr1.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr1[i]);
        }
        
        //method params
        let array = worksheet.GetRange("A1:A12");
        
        let func = Api.GetWorksheetFunction();
        let ans = func.QUARTILE(array, quart);
        
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
