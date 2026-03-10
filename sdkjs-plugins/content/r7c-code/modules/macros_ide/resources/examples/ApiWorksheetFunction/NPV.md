/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.NPV
 * 
 *  Демонстрация использования метода NPV класса ApiWorksheetFunction
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
        // This example shows how to return the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values).
        
        // How to get the net present value of an investment.
        
        // Use a function to get the net present value of an investment based on different parameters.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Rate");
        worksheet.GetRange("A2").SetValue(0.05);
        let values = ["Payment", -10000, 3000, 4500, 6000];
        
        for (let i = 0; i < values.length; i++) {
            worksheet.GetRange("B" + (i + 1)).SetValue(values[i]);
        }
        let range = worksheet.GetRange("B2:B5");
        worksheet.GetRange("B6").SetValue(func.NPV(0.05, range));
        
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
