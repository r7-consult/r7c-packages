/**
 * @file SetSize_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDrawing.SetSize
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the size of a drawing object.
 * It creates a shape, sets its size to 120x70 units, and then sets its position on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить размер объекта рисования.
 * Он создает фигуру, устанавливает ее размер на 120x70 единиц, а затем устанавливает ее положение на листе.
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
        // This example sets the size of the shape bounding box.
        
        // How to change the size of the drawing.
        
        // Resize a drawing by setting its width and height.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let drawing = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        drawing.SetSize(120 * 36000, 70 * 36000);
        drawing.SetPosition(0, 2 * 36000, 2, 3 * 36000);
        
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
