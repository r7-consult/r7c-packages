/**
 * @file GetRange_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetRange
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an object that represents the selected range of the sheet.
 * It sets a value in cell A2, gets the range A1:D5, and then sets its horizontal alignment to "center".
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект, представляющий выбранный диапазон листа.
 * Он устанавливает значение в ячейке A2, получает диапазон A1:D5, а затем устанавливает его горизонтальное выравнивание по центру.
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
        // This example shows how to get an object that represents the selected range of the sheet.
        
        // How to get a range using address.
        
        // Get range and set its horizontal alignment.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A2").SetValue("2");
        let range = worksheet.GetRange("A1:D5");
        range.SetAlignHorizontal("center");
        
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
