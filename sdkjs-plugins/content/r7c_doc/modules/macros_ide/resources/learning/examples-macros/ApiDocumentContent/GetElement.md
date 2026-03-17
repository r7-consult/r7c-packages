/**
 * @file GetElement_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDocumentContent.GetElement
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get an element by its position in the document content within a shape.
 * It creates a shape, gets its content, retrieves the first paragraph, sets its justification to center,
 * and adds text to it.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить элемент по его позиции в содержимом документа внутри фигуры.
 * Он создает фигуру, получает ее содержимое, извлекает первый абзац, устанавливает его выравнивание по центру,
 * и добавляет к нему текст.
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
        // This example shows how to get an element by its position in the document content.
        
        // How to get an element of the document content knowing its index position.
        
        // Get a document element then change its position and content.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetJc("center");
        paragraph.AddText("This is a paragraph with the text in it aligned by the center. ");
        paragraph.AddText("The justification is specified in the paragraph style. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        
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
