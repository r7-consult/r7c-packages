/**
 * @file GetCustomProperties_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetCustomProperties
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve and manage custom properties within a spreadsheet
 * using `Api.GetCustomProperties()`.
 * It adds various types of custom properties (string, number, date, boolean) to the document,
 * retrieves them, and then displays these properties within a shape added to the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получать и управлять пользовательскими свойствами в электронной таблице
 * с помощью `Api.GetCustomProperties()`.
 * Он добавляет различные типы пользовательских свойств (строка, число, дата, булево) к документу,
 * извлекает их, а затем отображает эти свойства внутри фигуры, добавленной на активный лист.
 *
 * @returns {ApiCustomProperties} An object representing the collection of custom properties for the document. (Объект, представляющий коллекцию пользовательских свойств документа.)
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        // Инициализация API R7 Office
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        const worksheet = Api.GetActiveSheet();
        
        // Get the custom properties collection
        // Получение коллекции пользовательских свойств
        const customProps = Api.GetCustomProperties();
        
        // Add various custom properties
        // Добавление различных пользовательских свойств
        customProps.Add("MyStringProperty", "Hello, Spreadsheet!");
        customProps.Add("MyNumberProperty", 123.450);
        customProps.Add("MyDateProperty", new Date("23 November 2023"));
        customProps.Add("MyBoolProperty", true);
        
        // Retrieve the custom properties
        // Извлечение пользовательских свойств
        let stringValue = customProps.Get("MyStringProperty");
        let numberValue = customProps.Get("MyNumberProperty");
        let dateValue = customProps.Get("MyDateProperty");
        let boolValue = customProps.Get("MyBoolProperty");
        
        // Create a solid fill and a no-fill stroke for the shape
        // Создание сплошной заливки и обводки без заливки для фигуры
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a rectangular shape to the worksheet
        // Добавление прямоугольной фигуры на лист
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 50 * 36000,
        	fill, stroke,
        	0, 0, 5, 0
        );
        
        // Get the content of the shape and its first paragraph
        // Получение содержимого фигуры и ее первого абзаца
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        
        // Add the custom property values to the shape's text content
        // Добавление значений пользовательских свойств в текстовое содержимое фигуры
        paragraph.AddText("Custom String Property: " + stringValue + "\n");
        paragraph.AddText("Custom Number Property: " + numberValue + "\n");
        paragraph.AddText("Custom Date Property: " + dateValue.toDateString() + "\n");
        paragraph.AddText("Custom Boolean Property: " + boolValue);
        
        // Success notification
        // Уведомление об успешном выполнении
        console.log('Macro executed successfully');
        
    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user in cell A1 if API is available
        // Опционально: Показать ошибку пользователю в ячейке A1, если API доступен
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
