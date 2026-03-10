/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FORECAST_ETS
 * 
 *  Демонстрация использования метода FORECAST_ETS класса ApiWorksheetFunction
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
        // This example shows how to calculate or predict a future value based on existing (historical) values by using the AAA version of the Exponential Smoothing (ETS) algorithm.
        
        // How to calculate or predict a future value based on existing (historical) values.
        
        // Use an AAA version of the Exponential Smoothing (ETS) algorithm to calculate or predict a future value based on existing values.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let dates = ["10/1/2017", "11/1/2017", "12/1/2017", "1/1/2018", "2/1/2018", "3/1/2018"];
        let numbers = [12558, 14356, 16345, 18678, 14227];
        
        for (let i = 0; i < dates.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(dates[i]);
        }
        for (let j = 0; j < numbers.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(numbers[j]);
        }
        
        worksheet.GetRange("A1").SetColumnWidth(15);
        let range1 = worksheet.GetRange("B1:B5");
        let range2 = worksheet.GetRange("A1:A5");
        worksheet.GetRange("B6").SetValue(func.FORECAST_ETS("3/1/2018", range1, range2, 0, 1, 1));
        
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
