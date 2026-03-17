/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.STDEVP
 * 
 *  Демонстрация использования метода STDEVP класса ApiWorksheetFunction
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
        // This example shows how to calculate standard deviation based on the entire population given as arguments (ignores logical values and text).
        
        // How to calculate standard deviation based on the entire population.
        
        // Use a function to get the standard deviation.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [1, 0, 0, 0, 4, 1, 0, 0, 2, 3, 6, 7, 6, 8, 10, 12];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        let func = Api.GetWorksheetFunction();
        let ans = func.STDEVP(1, 0, 0, 0, 4, 1, 0, 0, 2, 3, 6, 7, 6, 8, 10, 12); 
        
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
