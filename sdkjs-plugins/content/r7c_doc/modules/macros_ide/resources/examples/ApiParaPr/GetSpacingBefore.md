/**
 * @file GetSpacingBefore_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.GetSpacingBefore
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the spacing before value of a paragraph.
 * It creates a shape, adds text to the first paragraph, creates a second paragraph, sets its spacing before, and then displays the spacing value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение интервала перед абзацем.
 * Он создает фигуру, добавляет текст в первый абзац, создает второй абзац, устанавливает его интервал перед, а затем отображает значение интервала.
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
        // This example shows how to get the spacing before value of the current paragraph.
        
        // How to get spacing information which is before the paragraph.
        
        // Get two consecutive paragraphs add spacing between them then get the spacing before second one and display it in the worksheet. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is an example of setting a space before a paragraph.");
        paragraph.AddText("The second paragraph will have an offset of one inch from the top. ");
        paragraph.AddText("This is due to the fact that the second paragraph has this offset enabled.");
        let paragraph2 = Api.CreateParagraph();
        paragraph2.AddText("This is the second paragraph and it is one inch away from the first paragraph.");
        content.Push(paragraph2);
        let paraPr = paragraph2.GetParaPr();
        paraPr.SetSpacingBefore(1440);
        let spacingBefore = paraPr.GetSpacingBefore();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Spacing before: " + spacingBefore);
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
