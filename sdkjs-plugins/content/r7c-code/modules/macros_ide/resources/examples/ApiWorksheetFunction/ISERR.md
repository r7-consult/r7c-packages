/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISERR
 * 
 *  Демонстрация использования метода ISERR класса ApiWorksheetFunction
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
        // This example shows how to check whether a value is an error other than N/A, and returns true or false.
        
        // How to check if the cell contains an error and not N/A value.
        
        // Use a function to check whether the value is error or not and is not N/A.
        
        const worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("B3").SetValue("45")
        let result = func.ISERROR("B3");
        worksheet.GetRange("C3").SetValue(result)
        
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
