/**
 * @file SetFill_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetFill
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text color for the current text run.
 * It creates a shape, adds a run with normal text, and then adds another run with gray-colored text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить цвет текста для текущего текстового запуска.
 * Он создает фигуру, добавляет запуск с обычным текстом, а затем добавляет еще один запуск с текстом серого цвета.
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
        // This example sets the text color to the current text run.
        
        // How to color a text object.
        
        // Create a text run object, add a color to it using RGB format.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        fill = Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128));
        run.SetFill(fill);
        run.AddText("This is a text run with the font color set to gray.");
        paragraph.AddElement(run);
        
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
