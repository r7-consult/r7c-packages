/**
 * @file GetItalic_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.GetItalic
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the italic property of a text.
 * It creates a shape, adds a run with text, sets its italic property, and then displays whether it is italic.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство курсива текста.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его свойство курсива, а затем отображает, является ли он курсивом.
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
        // This example gets a text italic property.
        
        // How to find out whether a text is italic or not.
        
        // Get a text italic property as a boolean value.
        
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
        textProps.SetItalic(true);
        paragraph = Api.CreateParagraph();
        let isItalic = textProps.GetItalic();
        paragraph.AddText("Italic property: " + isItalic);
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
