/**
 * @file CreateBullet_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateBullet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a custom bullet for a paragraph.
 * It adds a shape to the active worksheet, then creates a bullet using a hyphen character,
 * applies this bullet to a paragraph, adds sample text to it, and finally adds this
 * paragraph to the shape's document content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать пользовательский маркер для абзаца.
 * Он добавляет фигуру на активный лист, затем создает маркер, используя символ дефиса,
 * применяет этот маркер к абзацу, добавляет к нему примерный текст, и, наконец, добавляет этот
 * абзац в содержимое документа фигуры.
 *
 * @param {string} sType - The type of bullet to use (e.g., "-", "•"). (Тип маркера для использования.)
 * @returns {ApiBullet} A new ApiBullet object. (Новый объект ApiBullet.)
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
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        
        // Get the document content of the shape
        // Получение содержимого документа фигуры
        let docContent = shape.GetContent();
        
        // Get the first paragraph of the shape's content
        // Получение первого абзаца содержимого фигуры
        let paragraph = docContent.GetElement(0);
        
        // Create a bullet using a hyphen
        // Создание маркера с использованием дефиса
        let bullet = Api.CreateBullet("-");
        
        // Set the bullet for the paragraph
        // Установка маркера для абзаца
        paragraph.SetBullet(bullet);
        
        // Add sample text to the paragraph
        // Добавление примерного текста в абзац
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        
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
