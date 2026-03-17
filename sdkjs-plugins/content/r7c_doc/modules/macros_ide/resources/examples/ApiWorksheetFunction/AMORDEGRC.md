/**
 * @file AMORDEGRC_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheetFunction.AMORDEGRC
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to return the prorated linear depreciation of an asset for each accounting period.
 * It calculates the prorated linear depreciation for an asset with specified cost, date purchased, first period, salvage, period, rate, and basis, then displays the result in cell A1.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как вернуть пропорциональную линейную амортизацию актива за каждый учетный период.
 * Он вычисляет пропорциональную линейную амортизацию для актива с указанной стоимостью, датой покупки, первым периодом, ликвидационной стоимостью, периодом, ставкой и базисом, а затем отображает результат в ячейке A1.
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
        // This example shows how to return the prorated linear depreciation of an asset for each accounting period.
        
        // How to get a prorated linear depreciation of an asset for each accounting period and display it in the worksheet.
        
        // Get a function that gets prorated linear depreciation of an asset for each accounting period.
        
        let worksheet = Api.GetActiveSheet();
        let func = Api.GetWorksheetFunction();
        worksheet.GetRange("A1").SetValue(func.AMORDEGRC(3500, "1/1/2018", "3/1/2018", 500, 1, 0.25, 1));
        
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
