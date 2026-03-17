/**
 * @file GetIndFirstLine_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.GetIndFirstLine
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the first line indentation of a paragraph.
 * It creates a shape, gets its content, sets the first line indentation of the paragraph properties, and then displays the indentation value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить отступ первой строки абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает отступ первой строки свойств абзаца, а затем отображает значение отступа.
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
        // This example shows how to get the paragraph first line indentation.
        
        // How to get first line indent of a paragraph.
        
        // Get paragraph lines using the indent order.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetIndFirstLine(1440);
        paragraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
        paragraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        let indFirstLine = paraPr.GetIndFirstLine();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("First line indent: " + indFirstLine);
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
