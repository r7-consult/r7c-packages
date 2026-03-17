/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FVSCHEDULE
 * 
 *  Демонстрация использования метода FVSCHEDULE класса ApiWorksheetFunction
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
        // This example shows how to return the future value of an initial principal after applying a series of compound interest rates.
        
        // How to get the future value of an initial principal.
        
        // Use a function to get future value of an initial principal based on different parameters.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Principal");
        worksheet.GetRange("B1").SetValue(100);
        worksheet.GetRange("A2").SetValue("Rate 1");
        worksheet.GetRange("B2").SetValue(0.03);
        worksheet.GetRange("A3").SetValue("Rate 2");
        worksheet.GetRange("B3").SetValue(0.05);
        worksheet.GetRange("A4").SetValue("Rate 3");
        worksheet.GetRange("B4").SetValue(0.1);
        let range = worksheet.GetRange("B2:B4");
        worksheet.GetRange("B5").SetValue(func.FVSCHEDULE(100, range));
        
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
