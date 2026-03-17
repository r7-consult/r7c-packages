/**
 * @file GetIndRight_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.GetIndRight
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the right side indentation of a paragraph.
 * It creates a shape, gets its content, sets the right indentation and justification of the paragraph properties, and then displays the indentation value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить отступ с правой стороны абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает правый отступ и выравнивание свойств абзаца, а затем отображает значение отступа.
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
        // This example shows how to get the paragraph right side indentation.
        
        // How to get a right indent of a paragraph.
        
        // Get the right paragraph indent by the side.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetIndRight(2880);
        paraPr.SetJc("right");
        paragraph.AddText("This is the first paragraph with the right offset of 2 inches set to it. ");
        paragraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
        let indRight = paraPr.GetIndRight();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Right indent: " + indRight);
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
