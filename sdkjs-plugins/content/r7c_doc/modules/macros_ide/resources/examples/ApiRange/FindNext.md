/**
 * @file FindNext_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.FindNext
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to find the next cell that matches specific conditions within a range.
 * It sets up a sample data table, finds the first occurrence of a value, and then finds and highlights the next occurrence.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как найти следующую ячейку, соответствующую определенным условиям в диапазоне.
 * Он настраивает образец таблицы данных, находит первое вхождение значения, а затем находит и выделяет следующее вхождение.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example finds the next cell that matches those same conditions.
        
        // How to get the next cell from a range that meets search requirements.
        
        // Get a range, find the next cell that satisfies search conditions and fill it with color.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(2014);
        worksheet.GetRange("C1").SetValue(2015);
        worksheet.GetRange("D1").SetValue(2016);
        worksheet.GetRange("A2").SetValue("Projected Revenue");
        worksheet.GetRange("A3").SetValue("Estimated Costs");
        worksheet.GetRange("A4").SetValue("Cost price");
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(250);
        worksheet.GetRange("B4").SetValue(50);
        worksheet.GetRange("C2").SetValue(200);
        worksheet.GetRange("C3").SetValue(260);
        worksheet.GetRange("C4").SetValue(120);
        worksheet.GetRange("D2").SetValue(200);
        worksheet.GetRange("D3").SetValue(200);
        worksheet.GetRange("D4").SetValue(160);
        let range = worksheet.GetRange("A2:D4");
        let searchRange = range.Find("200", "B1", "xlValues", "xlWhole", "xlByColumns", "xlNext", true);
        searchRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        let nextSearchRange = range.FindNext(searchRange);
        nextSearchRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191));
        
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
