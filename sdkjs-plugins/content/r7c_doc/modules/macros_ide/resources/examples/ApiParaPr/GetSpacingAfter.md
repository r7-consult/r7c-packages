/**
 * @file GetSpacingAfter_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiParaPr.GetSpacingAfter
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the spacing after value of a paragraph.
 * It creates a shape, gets its content, sets the spacing after the paragraph properties, and then displays the spacing value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение интервала после абзаца.
 * Он создает фигуру, получает ее содержимое, устанавливает интервал после свойств абзаца, а затем отображает значение интервала.
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
        // This example shows how to get the spacing after value of the current paragraph.
        
        // How to get spacing information which is after the paragraph.
        
        // Get two consecutive paragraphs add spacing between them then get the spacing after first one and display it in the worksheet. 
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let paraPr = paragraph.GetParaPr();
        paraPr.SetSpacingAfter(1440);
        paragraph.AddText("This is an example of setting a space after a paragraph. ");
        paragraph.AddText("The second paragraph will have an offset of one inch from the top. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
        paragraph.AddText("These sentences are used to add lines for demonstrative purposes.");
        let spacingAfter = paraPr.GetSpacingAfter();
        paragraph = Api.CreateParagraph();
        paragraph.AddText("Spacing after : " + spacingAfter);
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
