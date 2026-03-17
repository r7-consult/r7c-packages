/**
 * @file SetVerticalTextAlign_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiShape.SetVerticalTextAlign
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the vertical alignment to the shape content where a paragraph or text runs can be inserted.
 * It creates a shape, gets its content, removes all existing elements, sets the vertical text alignment to "bottom", creates a new paragraph with text, and then adds this paragraph to the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить вертикальное выравнивание содержимого фигуры, куда можно вставлять абзацы или текстовые фрагменты.
 * Он создает фигуру, получает ее содержимое, удаляет все существующие элементы, устанавливает вертикальное выравнивание текста на «bottom», создает новый абзац с текстом, а затем добавляет этот абзац в содержимое фигуры.
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
        // This example sets the vertical alignment to the shape content where a paragraph or text runs can be inserted.
        
        // How to specify a vertical alignment of a shape content.
        
        // Set text vertical alignment of a shape to bottom.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 50 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        content.RemoveAllElements();
        shape.SetVerticalTextAlign("bottom");
        let paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("We removed all elements from the shape and added a new paragraph inside it ");
        paragraph.AddText("aligning it vertically by the bottom.");
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
