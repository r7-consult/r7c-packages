/**
 * @file SetCreator_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCore.SetCreator
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the creator of the current workbook using the ApiCore.
 * It sets a creator for the workbook and then displays it in a shape on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить создателя текущей рабочей книги с помощью ApiCore.
 * Он устанавливает создателя для рабочей книги, а затем отображает его в фигуре на листе.
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
        // This example demonstrates how to set the creator of the current workbook using the ApiCore.
        
        const worksheet = Api.GetActiveSheet();
        const core = Api.GetCore();
        
        core.SetCreator("John Smith");
        const creator = core.GetCreator();
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(100, 50, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 100 * 36000,
        	fill, stroke,
        	0, 0, 3, 0
        );
        
        let paragraph = shape.GetContent().GetElement(0);
        paragraph.AddText("Creator: " + creator);
        
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
