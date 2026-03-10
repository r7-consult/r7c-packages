/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiTextPr.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of an ApiTextPr object.
 * It creates a shape, gets its content, creates a run, retrieves its text properties, and then displays the class type.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiTextPr.
 * Он создает фигуру, получает ее содержимое, создает запуск, извлекает его текстовые свойства, а затем отображает тип класса.
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
        // This example gets a class type and pastes it into the presentation.
        
        // How to get a class type of ApiTextPr.
        
        // Get a class type of ApiTextPr and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        let run = Api.CreateRun();
        let textProps = run.GetTextPr();
        textProps.SetFontSize(30);
        paragraph.SetJc("left");
        let classType = textProps.GetClassType();
        run.AddText("Class Type = " + classType);
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
