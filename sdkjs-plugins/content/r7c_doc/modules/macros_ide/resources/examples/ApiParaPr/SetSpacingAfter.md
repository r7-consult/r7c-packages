/**
 * @file SetSpacingAfter_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.SetSpacingAfter
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the spacing after a paragraph.
 * It creates a shape, gets its content, sets the spacing after the paragraph properties, and then adds text to the paragraph.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить интервал после абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает интервал после свойств абзаца, а затем добавляет текст в абзац.
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
        // This example sets the spacing after the current paragraph.
        
        // How to add the spacing after the paragraphs using points.
        
        // Get a paragraph from the shape's content then add a text specifying the spacing after a custom text.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetSpacingAfter(1440);
        paragraph.AddText("This is an example of setting a space after a paragraph. ");
        paragraph.AddText("The second paragraph will have an offset of one inch from the top. ");
        paragraph.AddText("This is due to the fact that the first paragraph has this offset enabled.");
        paragraph = Api.CreateParagraph();
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
