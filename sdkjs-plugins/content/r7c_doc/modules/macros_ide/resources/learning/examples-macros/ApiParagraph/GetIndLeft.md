/**
 * @file GetIndLeft_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetIndLeft
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the left side indentation of a paragraph.
 * It creates a shape, adds text to the first paragraph, sets its left indentation, and then displays the indentation value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить отступ с левой стороны абзаца.
 * Он создает фигуру, добавляет текст в первый абзац, устанавливает его левый отступ, а затем отображает значение отступа.
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
        // This example shows how to get the paragraph left side indentation.
        
        // How to get a left indent of a paragraph.
        
        // Get the left paragraph indent by the side.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("This is a paragraph with the indent of 2 inches set to it. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.SetIndLeft(2880);
        let indLeft = paragraph.GetIndLeft();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Left indent: " + indLeft);
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
