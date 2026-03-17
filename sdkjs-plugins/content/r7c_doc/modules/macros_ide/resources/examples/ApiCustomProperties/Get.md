/**
 * @file Get_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCustomProperties.Get
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the value of a custom property by its name.
 * It adds an existing custom property and then attempts to retrieve both an existing and a non-existent property,
 * displaying their values in a shape on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить значение пользовательского свойства по его имени.
 * Он добавляет существующее пользовательское свойство, а затем пытается получить как существующее, так и несуществующее свойство,
 * отображая их значения в фигуре на листе.
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
        // This example demonstrates how to get the value of a custom property by its name.
        
        const worksheet = Api.GetActiveSheet();
        const customProps = Api.GetCustomProperties();
        
        customProps.Add("ExistingProp", "#123456");
        
        const existingProp = customProps.Get("ExistingProp");
        const nonExistentProp = customProps.Get("NonExistentProp");
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 50 * 36000,
        	fill, stroke,
        	0, 0, 5, 0
        );
        
        let paragraph = shape.GetDocContent().GetElement(0);
        paragraph.AddText("Existing Property Value: " + existingProp);
        paragraph.AddText("\nNon-Existent Property Value: " + nonExistentProp);
        
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
