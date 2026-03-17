/**
 * @file GetRotation_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDrawing.GetRotation
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the rotation angle of a drawing object.
 * It creates a shape, sets its size, position, and rotation, and then displays the rotation angle.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить угол поворота объекта рисования.
 * Он создает фигуру, устанавливает ее размер, положение и поворот, а затем отображает угол поворота.
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
        // This example shows how to get the rotation angle of the drawing.
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let drawing = worksheet.AddShape("rect", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        drawing.SetSize(130 * 36000, 10 * 36000);
        drawing.SetPosition(0, 2 * 36000, 1, 5 * 36000);
        drawing.SetRotation(90);
        let rotAngle = drawing.GetRotation();
        worksheet.GetRange("A1").SetValue("Drawing rotation angle is: " + rotAngle + " degrees");
        
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
