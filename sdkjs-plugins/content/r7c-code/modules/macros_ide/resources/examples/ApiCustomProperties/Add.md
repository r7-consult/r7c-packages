/**
 * @file Add_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCustomProperties.Add
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to add various types of custom properties to a document.
 * It adds string, boolean, number (integer and real), and date custom properties,
 * then retrieves and displays them in a shape on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как добавлять различные типы пользовательских свойств в документ.
 * Он добавляет строковые, булевы, числовые (целые и вещественные) и датовые пользовательские свойства,
 * затем извлекает и отображает их в фигуре на листе.
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
        // This example demonstrates how to add a boolean custom property.
        
        const worksheet = Api.GetActiveSheet();
        const customProps = Api.GetCustomProperties();
        
        customProps.Add("CompanyName", "R7 Office");
        const companyName = customProps.Get("CompanyName");
        
        customProps.Add("TrueBool", true);
        const trueBool = customProps.Get("TrueBool");
        
        customProps.Add("NumberInt", 3.0);
        customProps.Add("NumberReal", 3.14);
        const numberInt = customProps.Get("NumberInt")
        const numberReal = customProps.Get("NumberReal")
        
        customProps.Add("BirthDate", new Date("20 January 2000"));
        const birthDate = customProps.Get("BirthDate");
        const isOfLegalAge = (new Date().getFullYear() - birthDate.getFullYear()) >= 18;
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 100 * 36000,
        	fill, stroke,
        	0, 0, 5, 0
        );
        
        let paragraph = shape.GetDocContent().GetElement(0);
        
        paragraph.AddText("CompanyName: " + companyName);
        paragraph.AddLineBreak();
        
        paragraph.AddText("\nTrueBool: " + trueBool);
        paragraph.AddLineBreak();
        
        paragraph.AddText("\nNumberInt: " + numberInt);
        paragraph.AddText("\nNumberReal: " + numberReal);
        paragraph.AddLineBreak();
        
        paragraph.AddText("\nBirthDate: " + birthDate.toDateString());
        paragraph.AddText("\nIs of legal age: " + isOfLegalAge);
        
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
