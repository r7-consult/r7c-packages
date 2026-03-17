/**
 * @file GetFormula_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetFormula
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the formula of a specified range.
 * It sets values in cells B1 and C1, sets a formula in cell A1, and then displays the formula from cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить формулу указанного диапазона.
 * Он устанавливает значения в ячейках B1 и C1, устанавливает формулу в ячейке A1, а затем отображает формулу из ячейки A1.
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
        // This example shows how to get a formula of the specified range.
        
        // How to find out a range formula.
        
        // Get a range, get its cell formula and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(1);
        worksheet.GetRange("C1").SetValue(2);
        let range = worksheet.GetRange("A1");
        range.SetValue("=SUM(B1:C1)");
        let formula = range.GetFormula();
        worksheet.GetRange("A3").SetValue("Formula from cell A1: " + formula);
        
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
