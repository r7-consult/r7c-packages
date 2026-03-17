/**
 * @file SetColor_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text color for the current text run in RGB format.
 * It creates a shape, adds a run with text, and then sets its color to gray.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить цвет текста для текущего текстового запуска в формате RGB.
 * Он создает фигуру, добавляет запуск с текстом, а затем устанавливает его цвет в серый.
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
        // This example sets the text color for the current text run in the RGB format.
        
        // How to change text color.
        
        // Create a text run object, update its font color using RGB format values.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is a text run with the font color set to gray.");
        paragraph.AddElement(run);
        run.SetColor(128, 128, 128);
        
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
