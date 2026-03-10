/**
 * @file GetJc_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.GetJc
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the justification of a paragraph's content.
 * It creates a shape, gets its content, sets the justification of the paragraph properties to center, and then displays the justification type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить выравнивание содержимого абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает выравнивание свойств абзаца по центру, а затем отображает тип выравнивания.
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
        // This example shows how to get the paragraph contents justification.
        
        // How to know the justification type of a paragraph.
        
        // Create a paragraph set justification type, get it and display in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetJc("center");
        paragraph.AddText("This is a paragraph with the text in it aligned by the center. ");
        paragraph.AddText("The justification is specified in the paragraph style. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        let justification = paraPr.GetJc();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Justification: " + justification);
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
