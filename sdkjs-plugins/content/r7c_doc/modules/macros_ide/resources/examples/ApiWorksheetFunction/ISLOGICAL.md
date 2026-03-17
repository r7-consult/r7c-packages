/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISLOGICAL
 * 
 *  Демонстрация использования метода ISLOGICAL класса ApiWorksheetFunction
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
        // This example shows how to check whether a value is a logical value (true or false), and returns true or false. 
        
        // How to check if the cell contains a logical value.
        
        // Use a function to check whether a range data is a logical value.
        
        const worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B3").SetValue("66");
        
        let func = Api.GetWorksheetFunction();
        let result = func.ISLOGICAL(worksheet.GetRange("B3"));
        worksheet.GetRange("C3").SetValue(result);
        
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
