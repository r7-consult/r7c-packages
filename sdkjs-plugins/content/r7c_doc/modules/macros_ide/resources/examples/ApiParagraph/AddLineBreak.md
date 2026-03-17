/**
 * @file AddLineBreak_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.AddLineBreak
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add a line break to a paragraph within a shape.
 * It creates a shape, gets its content, adds text to the first paragraph, adds a line break,
 * and then adds more text on the new line.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить разрыв строки в абзац внутри фигуры.
 * Он создает фигуру, получает ее содержимое, добавляет текст в первый абзац, добавляет разрыв строки,
 * а затем добавляет больше текста на новой строке.
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
        // This example adds a line break to the current position and starts the next element from a new line.
        
        // How to insert a text from a new line.
        
        // Get the paragraph from the shape and add two sentences separated by a linebreak.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.SetJc("left");
        paragraph.AddText("This is a text inside the shape aligned left.");
        paragraph.AddLineBreak();
        paragraph.AddText("This is a text after the line break.");
        
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
