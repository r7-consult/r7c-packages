/**
 * @file GetBold_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.GetBold
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the bold property of a text.
 * It creates a shape, adds a run with text, sets its bold property, and then displays whether it is bold.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить свойство жирности текста.
 * Он создает фигуру, добавляет запуск с текстом, устанавливает его свойство жирности, а затем отображает, является ли он жирным.
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
        // This example gets a text bold using its property.
        
        // How to find out whether a text is bold or not.
        
        // Get a text bold property.
        
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
        textProps.SetBold(true);
        paragraph = Api.CreateParagraph();
        let isBold = textProps.GetBold();
        paragraph.AddText("Bold property: " + isBold);
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
