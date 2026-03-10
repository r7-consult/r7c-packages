/**
 * @file AddShape_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.AddShape
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a shape to the sheet with the specified parameters.
 * It creates a linear gradient fill and a stroke, and then adds a flowchart shape to the worksheet with these properties.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить фигуру на лист с указанными параметрами.
 * Он создает линейную градиентную заливку и обводку, а затем добавляет фигуру блок-схемы на лист с этими свойствами.
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
        // This example adds a shape to the sheet with the parameters specified.
        
        // How to add a shape to the worksheet.
        
        // Insert a flowchart shape to the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let gradientStop1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gradientStop2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        let fill = Api.CreateLinearGradientFill([gradientStop1, gradientStop2], 5400000);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        
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
