/**
 * R7 Office JavaScript макрос - ApiWorksheetFunction.FREQUENCY
 * 
 *  Демонстрация использования метода FREQUENCY класса ApiWorksheetFunction
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
        // This example shows how to calculate how often values occur within a range of values and then returns the first value of the returned vertical array of numbers.
        
        // How to get frequency of first value from a range.
        
        // Use a function to count how often values occur within a range.
        
        const worksheet = Api.GetActiveSheet();
        
        // Configure function parameters
        let data_array = [78, 74, 13, 17, 60]; // Historical data_array
        let bins_array = [78, 56, 87, 0, 19]; // Corresponding bins_array in Excel serial number format
        
        // Set data_array and bins_array in cells
        for (let i = 0; i < data_array.length; i++) {
          worksheet.GetRange("A" + (i + 1)).SetValue(bins_array[i])
          worksheet.GetRange("B" + (i + 1)).SetValue(data_array[i]);
        }
        
        // Get the ranges for data_array and bins_array
        let data_arrayRange = worksheet.GetRange("A1:A5");
        let bins_arrayRange = worksheet.GetRange("B1:B5");
        
        // Get the worksheet function object
        let func = Api.GetWorksheetFunction();
        
        // Ensure the ranges are properly passed to the function
        let forecast = func.FREQUENCY(data_arrayRange, bins_arrayRange);
        
        // Print the forecast result
        worksheet.GetRange("D1").SetValue(forecast);
        
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
