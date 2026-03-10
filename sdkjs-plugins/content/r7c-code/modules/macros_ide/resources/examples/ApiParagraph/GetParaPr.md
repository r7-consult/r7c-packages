/**
 * @file GetParaPr_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetParaPr
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the paragraph properties and set spacing.
 * It creates a shape, gets its content, retrieves the first paragraph, sets its spacing after,
 * and then adds text to both paragraphs.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойства абзаца и установить интервал.
 * Он создает фигуру, получает ее содержимое, извлекает первый абзац, устанавливает его интервал после,
 * а затем добавляет текст в оба абзаца.
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
        // This example shows how to get the paragraph properties.
        
        // How to get properites of a paragraph and set the spacing.
        
        // Get the paragraph properites, change them, add a text and add the paragraph to the shape content.
        
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
