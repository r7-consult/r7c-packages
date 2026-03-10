/**
 * @file GetHeight_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDrawing.GetHeight
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the height of a drawing object.
 * It creates a shape, sets its size and position, and then displays the height of the drawing object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить высоту объекта рисования.
 * Он создает фигуру, устанавливает ее размер и положение, а затем отображает высоту объекта рисования.
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
        // This example shows how to get the height of the drawing.
        
        // How to know a height of a drawing.
        
        // Get a drawing's height and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let drawing = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        drawing.SetSize(120 * 36000, 70 * 36000);
        drawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);
        let height = drawing.GetHeight();
        worksheet.GetRange("A1").SetValue("Drawing height = " + height);
        
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
