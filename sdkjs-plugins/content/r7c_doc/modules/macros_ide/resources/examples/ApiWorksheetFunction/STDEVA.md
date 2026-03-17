/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.STDEVA
 * 
 *  Демонстрация использования метода STDEVA класса ApiWorksheetFunction
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
        // This example shows how to estimate standard deviation based on a sample, including logical values and text. Text and the false logical value have the value 0; the true logical value has the value 1.
        
        // How to estimate standard deviation based on a sample considering logical and text data types.
        
        // Use a function to get the standard deviation.
        
        const worksheet = Api.GetActiveSheet();
        
        let valueArr = [1, 0, 0, 0, "text", 1, 0, 0, 2, 3, true, false, 6, 8, 10, 12];
        
        // Place the numbers in cells
        for (let i = 0; i < valueArr.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(valueArr[i]);
        }
        
        let func = Api.GetWorksheetFunction();
        let ans = func.STDEVA(
          1,
          0,
          0,
          0,
          "text",
          1,
          0,
          0,
          2,
          3,
          true,
          false,
          6,
          8,
          10,
          12
        ); //includes logical values
        
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
