/**
 * @file GetElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get a paragraph element by its position.
 * It creates a shape, adds multiple runs to the first paragraph, and then makes the third run bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить элемент абзаца по его позиции.
 * Он создает фигуру, добавляет несколько запусков в первый абзац, а затем делает третий запуск жирным.
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
        // This example shows how to get a paragraph element using the position specified.
        
        // How to get an element of a paragraph using its index.
        
        // Find a paragraph element using its index and change its properties.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.RemoveAllElements();
        let run = Api.CreateRun();
        run.AddText("This is the text for the first text run. Do not forget a space at its end to separate from the second one. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.AddText("This is the text for the second run. We will set it bold afterwards. It also needs space at its end. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.AddText("This is the text for the third run. It ends the paragraph.");
        paragraph.AddElement(run);
        run = paragraph.GetElement(2);
        run.SetBold(true);
        
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
