/**
 * @file SetLanguage_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiRun.SetLanguage
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to specify the language which will be used to check spelling and grammar (if requested) when processing the contents of this text run.
 * It creates a shape, adds a run with text, and then sets its language to English (Canada).
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как указать язык, который будет использоваться для проверки орфографии и грамматики (при необходимости) при обработке содержимого этого текстового запуска.
 * Он создает фигуру, добавляет запуск с текстом, а затем устанавливает его язык на английский (Канада).
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
        // This example specifies the languages which will be used to check spelling and grammar (if requested) when processing the contents of this text run.
        
        // How to set a language to the text for grammar checking.
        
        // Create a text run object, change its language to English for grammar check.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let content = shape.GetContent();
        let paragraph = content.GetElement(0);
        let run = Api.CreateRun();
        run.AddText("This is a text run with the text language set to English (Canada).");
        run.SetLanguage("en-CA");
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
