/**
 * @file SetFontFamily_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetFontFamily
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the specified font family as the font name for a text run.
 * It creates a shape, adds a run with normal text, and then adds another run with the font family set to 'Comic Sans MS'.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить указанное семейство шрифтов в качестве имени шрифта для текстового запуска.
 * Он создает фигуру, добавляет запуск с обычным текстом, а затем добавляет еще один запуск с семейством шрифтов, установленным на «Comic Sans MS».
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
        // This example sets all 4 font slots with the specified font family.
        
        // How to change the text font family.
        
        // Create a text run object and set the "Comic Sans MS" font faminy to it.
        
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
        run.SetFontFamily("Comic Sans MS");
        run.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
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
