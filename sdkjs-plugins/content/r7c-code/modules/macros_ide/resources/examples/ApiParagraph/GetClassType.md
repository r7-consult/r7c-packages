/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParagraph.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of an ApiParagraph object.
 * It creates a shape, gets its content, retrieves the first paragraph, and then displays the class type of the paragraph object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiParagraph.
 * Он создает фигуру, получает ее содержимое, извлекает первый абзац, а затем отображает тип класса объекта абзаца.
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
        // This example gets a class type and inserts it into the document.
        
        // How to get a class type of ApiParagraph.
        
        // Get a class type of ApiParagraph and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let classType = paragraph.GetClassType();
        paragraph.AddText("Class Type = " + classType);
        
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
