/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.XNPV
 * 
 *  Демонстрация использования метода XNPV класса ApiWorksheetFunction
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
        // This example shows how to return the net present value for a schedule of cash flows.
        
        // How to return the net present value for a schedule of cash flows.
        
        // Use a function to return the net present value.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Rate");
        worksheet.GetRange("A2").SetValue(0.05);
        
        let payment = ["Payment/Income", -10000, 500, 5000, 3000];
        let dates = ["Payment dates", "1/1/2018", "4/1/2018", "8/1/2018", "12/1/2018"];
        
        for (let i = 0; i < payment.length; i++) {
            worksheet.GetRange("B" + (i + 1)).SetValue(payment[i]);
        }
        for (let j = 0; j < dates.length; j++) {
            worksheet.GetRange("C" + (j + 1)).SetValue(dates[j]);
        }
        
        worksheet.GetRange("C1").SetColumnWidth(15);
        let range1 = worksheet.GetRange("B2:B5");
        let range2 = worksheet.GetRange("C2:C5");
        worksheet.GetRange("D5").SetValue(func.XNPV(0.05, range1, range2));
        
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
