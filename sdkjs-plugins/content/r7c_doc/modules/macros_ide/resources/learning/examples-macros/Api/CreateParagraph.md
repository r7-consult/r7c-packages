/**
 * @file CreateParagraph_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateParagraph
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a new paragraph and add it to the content of a shape.
 * It adds a shape to the active worksheet, removes any existing content from it,
 * creates a new paragraph, sets its justification to left, adds sample text,
 * and then pushes this new paragraph into the shape's document content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать новый абзац и добавить его в содержимое фигуры.
 * Он добавляет фигуру на активный лист, удаляет из нее любое существующее содержимое,
 * создает новый абзац, устанавливает его выравнивание по левому краю, добавляет примерный текст,
 * а затем вставляет этот новый абзац в содержимое документа фигуры.
 *
 * @returns {ApiParagraph} A new empty ApiParagraph object. (Новый пустой объект ApiParagraph.)
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
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        
        // Get the document content of the shape and remove all existing elements
        // Получение содержимого документа фигуры и удаление всех существующих элементов
        let docContent = shape.GetContent();
        docContent.RemoveAllElements();
        
        // Create a new paragraph
        // Создание нового абзаца
        let paragraph = Api.CreateParagraph();
        
        // Set paragraph justification to left
        // Установка выравнивания абзаца по левому краю
        paragraph.SetJc("left");
        
        // Add sample text to the paragraph
        // Добавление примерного текста в абзац
        paragraph.AddText("We removed all elements from the shape and added a new paragraph inside it.");
        
        // Push the new paragraph into the shape's document content
        // Добавление нового абзаца в содержимое документа фигуры
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
