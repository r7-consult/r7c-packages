/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.TRANSPOSE
 * 
 *  Демонстрация использования метода TRANSPOSE класса ApiWorksheetFunction
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
        // This example shows how to convert a vertical range of cells to a horizontal range, or vice versa.
        
        // How to change orientation of cells to vertical/horizontal.
        
        // Use a function to transpose a range.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Ann");
        worksheet.GetRange("A2").SetValue("Bob");
        worksheet.GetRange("B1").SetValue("Apples");
        worksheet.GetRange("B2").SetValue("ranges");
        let range = worksheet.GetRange("A1:B2");
        worksheet.GetRange("A4:B5").SetValue(func.TRANSPOSE(range));
        
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
