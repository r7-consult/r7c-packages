/**
 * @file GetTextPr_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.GetTextPr
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the text properties of a run.
 * It creates a shape, gets its content, creates a run, retrieves its text properties, sets the font size, and then adds text to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текстовые свойства запуска.
 * Он создает фигуру, получает ее содержимое, создает запуск, извлекает его текстовые свойства, устанавливает размер шрифта, а затем добавляет текст в абзац.
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
        // This example shows how to get the text properties of the run.
        
        // How to change the text properties.
        
        // Create a text object, set its font size, justification, etc.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        let textProps = run.GetTextPr();
        textProps.SetFontSize(30);
        paragraph.SetJc("left");
        run.AddText("This is a sample text inside the shape with the font size set to 15 points using the text properties.");
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
