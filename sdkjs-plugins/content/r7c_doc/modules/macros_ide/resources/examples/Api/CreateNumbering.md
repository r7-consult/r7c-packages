/**
 * @file CreateNumbering_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateNumbering
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create numbered paragraphs within a shape.
 * It adds a shape to the active worksheet, then creates a numbering bullet style
 * ("ArabicParenR"). It applies this numbering to two paragraphs, adds sample text to them,
 * and pushes these paragraphs into the shape's document content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создавать нумерованные абзацы внутри фигуры.
 * Он добавляет фигуру на активный лист, затем создает стиль нумерованного маркера
 * ("ArabicParenR"). Он применяет эту нумерацию к двум абзацам, добавляет к ним примерный текст,
 * и вставляет эти абзацы в содержимое документа фигуры.
 *
 * @param {string} sType - The type of numbering to use (e.g., "ArabicParenR", "bullet"). (Тип нумерации для использования.)
 * @param {number} nStart - The starting number for the numbering. (Начальный номер для нумерации.)
 * @returns {ApiBullet} A new ApiBullet object representing the numbering. (Новый объект ApiBullet, представляющий нумерацию.)
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
        
        // Create a numbering bullet style
        // Создание стиля нумерованного маркера
        let bullet = Api.CreateNumbering("ArabicParenR", 1);
        
        // Create the first paragraph, apply numbering, and add text
        // Создание первого абзаца, применение нумерации и добавление текста
        let paragraph1 = Api.CreateParagraph();
        paragraph1.SetBullet(bullet);
        paragraph1.AddText(" This is an example of the numbered paragraph.");
        docContent.Push(paragraph1);
        
        // Create the second paragraph, apply numbering, and add text
        // Создание второго абзаца, применение нумерации и добавление текста
        let paragraph2 = Api.CreateParagraph();
        paragraph2.SetBullet(bullet);
        paragraph2.AddText(" This is an example of the numbered paragraph.");
        docContent.Push(paragraph2);
        
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
