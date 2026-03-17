/**
 * @file AddText_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.AddText
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add text to a paragraph within a shape.
 * It creates a shape, gets its content, sets the justification of the first paragraph to left,
 * and then adds two lines of text to it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавить текст в абзац внутри фигуры.
 * Он создает фигуру, получает ее содержимое, устанавливает выравнивание первого абзаца по левому краю,
 * а затем добавляет две строки текста в него.
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
        // This example adds some text to the paragraph.
        
        // How to add raw text to the paragraph.
        
        // Change content of the shape by adding a text.
        
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
