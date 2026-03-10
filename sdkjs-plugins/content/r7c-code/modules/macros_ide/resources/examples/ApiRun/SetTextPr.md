/**
 * @file SetTextPr_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetTextPr
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the text properties to the current run.
 * It creates a shape, adds a run with text, retrieves its text properties, sets the font size and bold property, and then applies these properties to the run.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить текстовые свойства для текущего запуска.
 * Он создает фигуру, добавляет запуск с текстом, извлекает его текстовые свойства, устанавливает размер шрифта и свойство жирности, а затем применяет эти свойства к запуску.
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
        // This example sets the text properties to the current run.
        
        // How to create text property for a text object.
        
        // Create a text run object, add properties like font size, style, color, etc.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is a sample text with the font size set to 15 points and the font weight set to bold.");
        let textProps = run.GetTextPr();
        textProps.SetFontSize(30);
        textProps.SetBold(true);
        run.SetTextPr(textProps);
        paragraph.AddElement(run);
        
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
