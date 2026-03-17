/**
 * @file COUNTIFS_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.COUNTIFS
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to count a number of cells specified by a given set of conditions or criteria.
 * It sets up data for buyers, products, and quantity, and then counts the number of cells where the product contains "apples" and the quantity is "45", displaying the result in cell E6.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как подсчитать количество ячеек, указанных заданным набором условий или критериев.
 * Он настраивает данные для покупателей, продуктов и количества, а затем подсчитывает количество ячеек, где продукт содержит «яблоки», а количество равно «45», отображая результат в ячейке E6.
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
        // This example shows how to count a number of cells specified by a given set of conditions or criteria.
        
        // How to find a number of cells that satisfy a list of conditions.
        
        // Use function to get cells if conditions are met.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        let buyer = ["Buyer", "Tom", "Bob", "Ann", "Kate", "John"];
        let product = ["Product", "Apples", "Red apples", "ranges", "Green apples", "ranges"];
        let quantity = ["Quantity", 12, 45, 18, 26, 10];
        
        for (let i = 0; i < buyer.length; i++) {
            worksheet.GetRange("A" + (i + 1)).SetValue(buyer[i]);
        }
        for (let j = 0; j < product.length; j++) {
            worksheet.GetRange("B" + (j + 1)).SetValue(product[j]);
        }
        for (let n = 0; n < quantity.length; n++) {
            worksheet.GetRange("C" + (n + 1)).SetValue(quantity[n]);
        }
        
        let range1 = worksheet.GetRange("B2:B6");
        let range2 = worksheet.GetRange("C2:C6");
        worksheet.GetRange("E6").SetValue(func.COUNTIFS(range1, "*apples", range2, "45"));
        
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
