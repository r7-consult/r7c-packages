/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiBullet.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the class type of an `ApiBullet` object.
 * It adds a shape to the active worksheet, creates a numbered bullet, applies it to paragraphs,
 * and then gets the class type of the bullet object, displaying it within the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта `ApiBullet`.
 * Он добавляет фигуру на активный лист, создает нумерованный маркер, применяет его к абзацам,
 * а затем получает тип класса объекта маркера, отображая его в содержимом фигуры.
 *
 * @returns {string} The class type of the ApiBullet object. (Тип класса объекта ApiBullet.)
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
        
        // Create a numbering bullet
        // Создание нумерованного маркера
        let bullet = Api.CreateNumbering("ArabicParenR", 1);
        
        // Set the bullet for the paragraph and add text
        // Установка маркера для абзаца и добавление текста
        paragraph.SetBullet(bullet);
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        
        // Create another paragraph, set the same bullet, and add text
        // Создание другого абзаца, установка того же маркера и добавление текста
        paragraph = Api.CreateParagraph();
        paragraph.SetBullet(bullet);
        paragraph.AddText(" This is an example of the bulleted paragraph.");
        docContent.Push(paragraph);
        
        // Get the class type of the bullet object
        // Получение типа класса объекта маркера
        let classType = bullet.GetClassType();
        
        // Create a new paragraph to display the class type
        // Создание нового абзаца для отображения типа класса
        paragraph = Api.CreateParagraph();
        paragraph.SetJc("left");
        paragraph.AddText("Class Type = " + classType);
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
