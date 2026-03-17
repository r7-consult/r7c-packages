/**
 * @file GetHidden_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRange.GetHidden
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the hidden property of a range.
 * It hides rows 1 to 3, sets values in cells A1:C1, and then displays whether the range is hidden.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить скрытое свойство диапазона.
 * Он скрывает строки с 1 по 3, устанавливает значения в ячейках A1:C1, а затем отображает, скрыт ли диапазон.
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
        // This example shows how to get the value hiding property.
        
        // How to find out hidden property from a range.
        
        // Get a range, get its cell hiding property and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let range = worksheet.GetRows("1:3");
        range.SetHidden(true);
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        let hidden = range.GetHidden();
        worksheet.GetRange("A4").SetValue("The values from A1:C1 are hidden: " + hidden);
        
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
