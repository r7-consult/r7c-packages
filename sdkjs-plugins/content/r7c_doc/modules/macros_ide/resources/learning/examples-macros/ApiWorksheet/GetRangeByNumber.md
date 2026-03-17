/**
 * @file GetRangeByNumber_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetRangeByNumber
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an object that represents the selected range of the sheet using the row/column coordinates for the cell selection.
 * It gets a range by its row and column numbers and then sets its value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить объект, представляющий выбранный диапазон листа, используя координаты строки/столбца для выбора ячейки.
 * Он получает диапазон по номерам строк и столбцов, а затем устанавливает его значение.
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
        // This example shows how to get an object that represents the selected range of the sheet using the row/column coordinates for the cell selection.
        
        // How to get a range using its coordinates.
        
        // Get range by number and set its value.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRangeByNumber(1, 2).SetValue("42");
        
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
