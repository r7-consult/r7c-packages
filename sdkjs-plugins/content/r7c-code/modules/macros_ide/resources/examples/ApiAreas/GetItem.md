/**
 * @file GetItem_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiAreas.GetItem
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve a single object from a collection of areas by its index.
 * It sets a value in a range, selects it, gets the areas collection, and then retrieves the item
 * at index 1 (the second item) from this collection. Finally, it displays information about the
 * retrieved item in the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить один объект из коллекции областей по его индексу.
 * Он устанавливает значение в диапазоне, выбирает его, получает коллекцию областей, а затем извлекает элемент
 * по индексу 1 (второй элемент) из этой коллекции. Наконец, он отображает информацию о
 * полученном элементе на листе.
 *
 * @param {number} index - The zero-based index of the item to retrieve from the collection. (Нулевой индекс элемента для извлечения из коллекции.)
 * @returns {ApiRange} The ApiRange object at the specified index. (Объект ApiRange по указанному индексу.)
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
        
        // Set a value in range B1:D1 and select it
        // Установка значения в диапазоне B1:D1 и его выбор
        let range = worksheet.GetRange("B1:D1");
        range.SetValue("1");
        range.Select();
        
        // Get the areas collection from the selected range
        // Получение коллекции областей из выбранного диапазона
        let areas = range.GetAreas();
        
        // Get the item at index 1 (second item) from the areas collection
        // Получение элемента по индексу 1 (второй элемент) из коллекции областей
        let item = areas.GetItem(1);
        
        // Display information about the retrieved item in the worksheet
        // Отображение информации о полученном элементе на листе
        range = worksheet.GetRange('A5');
        range.SetValue("The first item from the areas: ");
        range.AutoFit(false, true);
        worksheet.GetRange('B5').Paste(item);
        
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
