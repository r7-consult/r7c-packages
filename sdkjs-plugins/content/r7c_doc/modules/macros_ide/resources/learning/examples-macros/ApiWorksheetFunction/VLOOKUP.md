/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.VLOOKUP
 * 
 *  Демонстрация использования метода VLOOKUP класса ApiWorksheetFunction
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
        // This example shows how to look for a value in the leftmost column of a table and then returns a value in the same row from the specified column. By default, the table must be sorted in an ascending order.
        
        // How to look for a value in the leftmost column of a table.
        
        // Use a find a value in the leftmost column of a table and display it in the row.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let ids = ["ID", 1, 2, 3, 4, 5];
        let clients = ["Client", "John Smith", "Ella Tompson", "Mary Shinoda", "Lily-Ann Bates", "Clara Ray"];
        let phones = ["Phone number", "12054097166", "13343943678", "12568542099", "12057032298", "12052914781"];
        
        for (let i = 0; i < ids.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(ids[i]);
        }
        for (let j = 0; j < clients.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(clients[j]);
        }
        for (let n = 0; n < phones.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(phones[n]);
        }
        
        let range = worksheet.GetRange("A1:C5");
        worksheet.GetRange("D6").SetValue(func.VLOOKUP(3, range, 2, true));
        
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
