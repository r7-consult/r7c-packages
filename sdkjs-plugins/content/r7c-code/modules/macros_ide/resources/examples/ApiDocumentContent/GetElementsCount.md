/**
 * @file GetElementsCount_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDocumentContent.GetElementsCount
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the number of elements in the document content within a shape.
 * It creates a shape, gets its content, adds text to the first paragraph, and then displays the count of elements.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить количество элементов в содержимом документа внутри фигуры.
 * Он создает фигуру, получает ее содержимое, добавляет текст в первый абзац, а затем отображает количество элементов.
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
        // This example shows how to get a number of elements in the current document content.
        
        // How to get a number of elements of a shape from a document content.
        
        // Get a shape than count number of elements and display it using paragraph.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("We got the first paragraph inside the shape.");
        paragraph.AddLineBreak();
        paragraph.AddText("Number of elements inside the shape: " + content.GetElementsCount());
        paragraph.AddLineBreak();
        paragraph.AddText("Line breaks are NOT counted into the number of elements.");
        
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
