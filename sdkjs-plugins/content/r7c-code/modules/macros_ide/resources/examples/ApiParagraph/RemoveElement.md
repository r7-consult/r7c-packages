/**
 * @file RemoveElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.RemoveElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to remove an element from a paragraph by its position.
 * It creates a shape, adds multiple runs to the first paragraph, removes the third run, and then adds more runs.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить элемент из абзаца по его позиции.
 * Он создает фигуру, добавляет несколько запусков в первый абзац, удаляет третий запуск, а затем добавляет еще запусков.
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
        // This example removes an element using the position specified.
        
        // How to delete a paragraph element knowing its index.
        
        // Change the content of a shape by removing elements.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.RemoveAllElements();
        let run = Api.CreateRun();
        run.AddText("This is the first paragraph element. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.AddText("This is the second paragraph element. ");
        paragraph.AddElement(run);
        run = Api.CreateRun();
        run.AddText("This is the third paragraph element (it will be removed from the paragraph and we will not see it). ");
        paragraph.AddElement(run);
        paragraph.AddLineBreak();
        run = Api.CreateRun();
        run.AddText("This is the fourth paragraph element - it became the third, because we removed the previous run from the paragraph. ");
        paragraph.AddElement(run);
        paragraph.AddLineBreak();
        run = Api.CreateRun();
        run.AddText("Please note that line breaks are not counted into paragraph elements!");
        paragraph.AddElement(run);
        paragraph.RemoveElement(3);
        
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
