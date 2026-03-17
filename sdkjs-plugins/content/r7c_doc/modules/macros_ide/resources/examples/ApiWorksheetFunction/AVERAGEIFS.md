/**
 * @file AVERAGEIFS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AVERAGEIFS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to find the average (arithmetic mean) for the cells specified by a given set of conditions or criteria.
 * It sets up data for year, products, and income, then calculates the average income for products containing "Apples" in the year 2016, and displays the result in cell E6.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как найти среднее (арифметическое) для ячеек, указанных заданным набором условий или критериев.
 * Он настраивает данные для года, продуктов и дохода, затем вычисляет средний доход для продуктов, содержащих «Яблоки» в 2016 году, и отображает результат в ячейке E6.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example shows how to find the average (arithmetic mean) for the cells specified by a given set of conditions or criteria.
        
        // How to find an average if list of conditions are met.
        
        // Use function to get an average (arithmetic mean) of the cells if the set of requirements is satisfied.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let year = [2016, 2016, 2016, 2017, 2017, 2017];
        let products = ["Apples", "Red apples", "ranges", "Green apples", "Apples", "Bananas"];
        let income = ["$100.00", "$150.00", "$250.00", "$50.00", "$150.00", "$200.00"];
        
        for (let i = 0; i < year.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(year[i]);
        }
        for (let j = 0; j < products.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(products[j]);
        }
        for (let n = 0; n < income.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(income[n]);
        }
        
        let range1 = worksheet.GetRange("C1:C6");
        let range2 = worksheet.GetRange("B1:B6");
        let range3 = worksheet.GetRange("A1:A6");
        worksheet.GetRange("E6").SetValue(func.AVERAGEIFS(range1, range2, "*Apples", range3, 2016));
        
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
