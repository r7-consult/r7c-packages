/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.XIRR
 * 
 *  Демонстрация использования метода XIRR класса ApiWorksheetFunction
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
        // This example shows how to return the internal rate of return for a schedule of cash flows.
        
        // How to return the internal rate of return.
        
        // Use a function to return the internal rate of return.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let values = ["Values", "-$40,000.00", "$10,000.00", "$15,000.00", "$20,000.00"];
        let dates = ["Dates", "1/1/2018", "4/1/2018", "8/1/2018", "12/1/2018"];
        
        for (let i = 0; i < values.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(values[i]);
        }
        for (let j = 0; j < dates.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(dates[j]);
        }
        
        worksheet.GetRange("B1").SetColumnWidth(15);
        let range1 = worksheet.GetRange("A2:A5");
        let range2 = worksheet.GetRange("B2:B5");
        worksheet.GetRange("C5").SetValue(func.XIRR(range1, range2));
        
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
