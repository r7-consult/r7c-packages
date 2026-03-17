/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.DGET
 * 
 *  Демонстрация использования метода DGET класса ApiWorksheetFunction
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
        // This example shows how to extract from a database a single record that matches the conditions you specify.
        
        // How to get a value from a database range that satisfies a condition.
        
        // Use function to get a record after applying condition.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue("Name");
        worksheet.GetRange("B1").SetValue("Month");
        worksheet.GetRange("C1").SetValue("Sales");
        worksheet.GetRange("A2").SetValue("Alice");
        worksheet.GetRange("B2").SetValue("Jan");
        worksheet.GetRange("C2").SetValue(200);
        worksheet.GetRange("A3").SetValue("Alice");
        worksheet.GetRange("B3").SetValue("Feb");
        worksheet.GetRange("C3").SetValue(250);
        worksheet.GetRange("E1").SetValue("Name");
        worksheet.GetRange("E2").SetValue("Alice");
        worksheet.GetRange("F1").SetValue("Month");
        worksheet.GetRange("F2").SetValue("Feb");
        let range1 = worksheet.GetRange("A1:C3");
        let range2 = worksheet.GetRange("E1:F2");
        worksheet.GetRange("F4").SetValue(func.DGET(range1, "Sales", range2));
        
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
