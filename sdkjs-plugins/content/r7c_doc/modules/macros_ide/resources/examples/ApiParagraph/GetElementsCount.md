/**
 * @file GetElementsCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetElementsCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number of elements in a paragraph.
 * It creates a shape, gets its content, removes all elements from the first paragraph, and then adds a run with text and the count of elements.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество элементов в абзаце.
 * Он создает фигуру, получает ее содержимое, удаляет все элементы из первого абзаца, а затем добавляет запуск с текстом и количеством элементов.
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
        // This example shows how to get a number of elements in the current paragraph.
        
        // Get paragraph elements count.
        
        // How to get number of elements of the paragraph and show it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.RemoveAllElements();
        let run = Api.CreateRun();
        run.AddText("Number of paragraph elements at this point: ");
        run.AddTabStop();
        run.AddText("" + paragraph.GetElementsCount());
        run.AddLineBreak();
        paragraph.AddElement(run);
        run.AddText("Number of paragraph elements after we added a text run: ");
        run.AddTabStop();
        run.AddText("" + paragraph.GetElementsCount());
        
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
