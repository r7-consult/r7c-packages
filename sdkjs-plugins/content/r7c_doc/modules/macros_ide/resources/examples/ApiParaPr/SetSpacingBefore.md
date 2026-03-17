/**
 * @file SetSpacingBefore_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.SetSpacingBefore
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the spacing before a paragraph.
 * It creates a shape, adds text to the first paragraph, creates a second paragraph, sets its spacing before, and then adds it to the content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить интервал перед абзацем.
 * Он создает фигуру, добавляет текст в первый абзац, создает второй абзац, устанавливает его интервал перед, а затем добавляет его в содержимое.
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
        // This example sets the spacing before the current paragraph.
        
        // How to add the spacing before the paragraphs using points.
        
        // Get a paragraph from the shape's content then add a text specifying the spacing before a custom text.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is an example of setting a space before a paragraph. ");
        paragraph.AddText("The second paragraph will have an offset of one inch from the top. ");
        paragraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");
        paragraph = Api.CreateParagraph();
        let paraPr = paragraph.GetParaPr();
        paraPr.SetSpacingBefore(1440);
        paragraph.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
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
