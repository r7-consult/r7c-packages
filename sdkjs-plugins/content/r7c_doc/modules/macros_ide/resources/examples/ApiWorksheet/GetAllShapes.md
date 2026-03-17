/**
 * @file GetAllShapes_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllShapes
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all shapes from the sheet.
 * It adds a shape to the worksheet, and then retrieves all shapes to modify the first one by clearing its content, setting vertical text alignment, and adding a new paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все фигуры с листа.
 * Он добавляет фигуру на лист, а затем извлекает все фигуры, чтобы изменить первую, очистив ее содержимое, установив вертикальное выравнивание текста и добавив новый абзац.
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
        // This example shows how to get all shapes from the sheet.
        
        // How to get all shapes.
        
        // Get all shapes as an array.
        
        let worksheet = Api.GetActiveSheet();
        let gradientStop1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gradientStop2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        let fill = Api.CreateLinearGradientFill([gradientStop1, gradientStop2], 5400000);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let shapes = worksheet.GetAllShapes();
        let content = shapes[0].GetContent();
        content.RemoveAllElements();
        shapes[0].SetVerticalTextAlign("bottom");
        let paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("We removed all elements from the shape and added a new paragraph inside it ");
        paragraph.AddText("aligning it vertically by the bottom.");
        content.Push(paragraph);
        
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
