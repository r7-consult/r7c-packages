/**
 * @file SetSpacingLine_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.SetSpacingLine
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the line spacing for a paragraph.
 * It creates a shape, adds text to the first paragraph, sets its line spacing to auto, and then adds another paragraph with exact line spacing.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить межстрочный интервал для абзаца.
 * Он создает фигуру, добавляет текст в первый абзац, устанавливает его межстрочный интервал на авто, а затем добавляет еще один абзац с точным межстрочным интервалом.
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
        // This example sets the paragraph line spacing.
        
        // How to add a spacing line between paragraphs.
        
        // Get a paragraph from the shape's content then add a text specifying the spacing between text lines.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.SetSpacingLine(2 * 240, "auto");
        paragraph.AddText("Paragraph 1. Spacing: 2 times of a common paragraph line spacing.");
        paragraph.AddLineBreak();
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph = Api.CreateParagraph();
        paragraph.SetSpacingLine(200, "exact");
        paragraph.AddText("Paragraph 2. Spacing: exact 10 points.");
        paragraph.AddLineBreak();
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
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
