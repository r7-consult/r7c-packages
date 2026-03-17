/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.IRR
 * 
 *  Демонстрация использования метода IRR класса ApiWorksheetFunction
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
        // This example shows how to return the internal rate of return for a series of cash flows.
        
        // How to calculate the internal rate of the return for a series of cash flows.
        
        // Use a function to get the internal rate.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let values = ["Values", "-$40,000.00", "$10,000.00", "$15,000.00", "$20,000.00"];
        
        for (let i = 0; i < values.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(values[i]);
        }
        
        let range = worksheet.GetRange("A2:A5");
        worksheet.GetRange("B5").SetValue(func.IRR(range));
        
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
