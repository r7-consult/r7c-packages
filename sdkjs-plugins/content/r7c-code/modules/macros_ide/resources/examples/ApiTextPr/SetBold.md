/**
 * @file SetBold_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.SetBold
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the bold property to the text character.
 * It creates a shape, adds a run with text, sets its bold property, and then adds the run to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить свойство жирности для текстового символа.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его свойство жирности, а затем добавляет запуск в абзац.
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
        // This example sets the bold property to the text character.
        
        // How to make text bold.
        
        // Set bold property of a text.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        let textProps = run.GetTextPr();
        textProps.SetFontSize(30);
        textProps.SetBold(true);
        paragraph.SetJc("left");
        run.AddText("This is a sample text inside the shape with the font weight set to bold using the text properties.");
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
