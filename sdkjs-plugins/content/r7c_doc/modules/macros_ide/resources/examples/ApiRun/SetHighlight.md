/**
 * @file SetHighlight_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetHighlight
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify a highlighting color which is applied as a background to the contents of the run.
 * It creates a shape, adds a run with normal text, and then adds another run with light gray highlighted text.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать цвет выделения, который применяется в качестве фона к содержимому запуска.
 * Он создает фигуру, добавляет запуск с обычным текстом, а затем добавляет еще один запуск с текстом, выделенным светло-серым цветом.
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
        // This example specifies a highlighting color which is applied as a background to the contents of the run.
        
        // How to color a text object background.
        
        // Create a text run object, add color to its background.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.SetHighlight("lightGray");
        run.AddText("This is a text run with the text highlighted with light gray color.");
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
