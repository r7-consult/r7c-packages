/**
 * @file CreateTextPr_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateTextPr
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create and apply text properties to text within a shape.
 * It creates an empty `ApiTextPr` object, sets its font size to 30 and makes the text bold.
 * Then, it creates a paragraph, adds sample text, applies the defined text properties to it,
 * and inserts this paragraph into a newly added shape on the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создавать и применять свойства текста к тексту внутри фигуры.
 * Он создает пустой объект `ApiTextPr`, устанавливает его размер шрифта на 30 и делает текст жирным.
 * Затем он создает абзац, добавляет примерный текст, применяет к нему определенные свойства текста,
 * и вставляет этот абзац во вновь добавленную фигуру на активном листе.
 *
 * @returns {ApiTextPr} An empty ApiTextPr object. (Пустой объект ApiTextPr.)
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
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Create fill and stroke for the shape
        // Создание заливки и обводки для фигуры
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a shape to the worksheet
        // Добавление фигуры на лист
        let shape = worksheet.AddShape("flowChartOnlineStorage", 80 * 36000, 50 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        
        // Get the document content of the shape and remove all existing elements
        // Получение содержимого документа фигуры и удаление всех существующих элементов
        let docContent = shape.GetContent();
        docContent.RemoveAllElements();
        
        // Create text properties and set font size and bold
        // Создание свойств текста и установка размера шрифта и жирного начертания
        let textPr = Api.CreateTextPr();
        textPr.SetFontSize(30);
        textPr.SetBold(true);
        
        // Create a paragraph, set its justification, add text, and apply text properties
        // Создание абзаца, установка его выравнивания, добавление текста и применение свойств текста
        let paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("This is a sample text with the font size set to 30 and the font weight set to bold.");
        paragraph.SetTextPr(textPr);
        
        // Push the paragraph to the shape's document content
        // Добавление абзаца в содержимое документа фигуры
        docContent.Push(paragraph);
        
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
