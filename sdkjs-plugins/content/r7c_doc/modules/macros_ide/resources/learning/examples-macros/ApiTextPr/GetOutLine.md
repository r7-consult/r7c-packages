/**
 * @file GetOutLine_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.GetOutLine
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the outline property of a text.
 * It creates a shape, adds a run with text, sets its outline property, and then displays the outline type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство контура текста.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его свойство контура, а затем отображает тип контура.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
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
        // This example gets an outline property of a text.
        
        // How to find out whether a text is outlined or not.
        
        // Get a text outline property type, for example as a stroke.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is just a sample text. ");
        run.AddText("The text properties are changed and the style is added to the paragraph. ");
        run.AddLineBreak();
        paragraph.AddElement(run);
        let textProps = run.GetTextPr();
        stroke = Api.CreateStroke(0.2 * 36000, Api.CreateSolidFill(Api.CreateRGBColor(128, 128, 128)));
        textProps.SetOutLine(stroke);
        paragraph = Api.CreateParagraph();
        stroke = textProps.GetOutLine();
        let type = stroke.GetClassType();
        paragraph.AddText("Text outline type: " + type);
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
