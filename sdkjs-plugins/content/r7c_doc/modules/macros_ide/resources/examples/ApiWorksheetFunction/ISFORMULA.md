/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.ISFORMULA
 * 
 *  Демонстрация использования метода ISFORMULA класса ApiWorksheetFunction
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
        // This example shows how to check whether a reference to a cell contains a formula, and returns true or false. 
        
        // How to check if the cell contains formula or not.
        
        // Use a function to check whether a range data is a formula or not.
        
        const worksheet = Api.GetActiveSheet();
        
        // Set the formula in cell B3
        worksheet.GetRange("B3").SetValue("=SUM(5, 6)");
        
        // Check if there is a formula in C3
        let func = Api.GetWorksheetFunction();
        let result = func.ISFORMULA(worksheet.GetRange("B3"));
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
