/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.DCOUNT
 * 
 *  Демонстрация использования метода DCOUNT класса ApiWorksheetFunction
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
        // This example shows how to count the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
        
        // How to count the cells containing numbers in the field (column) of records in the database that match the conditions you specify.
        
        // Use function to count numbers from database records that met a condition specified.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Name");
        worksheet.GetRange("B1").SetValue("Age");
        worksheet.GetRange("C1").SetValue("Sales");
        worksheet.GetRange("A2").SetValue("Alice");
        worksheet.GetRange("B2").SetValue(20);
        worksheet.GetRange("C2").SetValue(200);
        worksheet.GetRange("A3").SetValue("Andrew");
        worksheet.GetRange("B3").SetValue(21);
        worksheet.GetRange("C3").SetValue(300);
        worksheet.GetRange("E1").SetValue("Sales");
        worksheet.GetRange("E2").SetValue(">200");
        let range1 = worksheet.GetRange("A1:C3");
        let range2 = worksheet.GetRange("E1:E2");
        worksheet.GetRange("E4").SetValue(func.DCOUNT(range1, "Sales", range2));
        
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
