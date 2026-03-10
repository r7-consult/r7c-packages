/**
 * @file GetLocale_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetLocale
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve the current locale ID of the document.
 * It first sets the locale to "en-CA" (English - Canada) for demonstration purposes,
 * then retrieves the current locale, and displays it in cell A1 of the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить текущий идентификатор локали документа.
 * Сначала он устанавливает локаль на "en-CA" (английский - Канада) для демонстрационных целей,
 * затем извлекает текущую локаль и отображает ее в ячейке A1 активного листа.
 *
 * @returns {string} The current locale ID (e.g., "en-US", "ru-RU"). (Текущий идентификатор локали.)
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
        
        // Set the locale to "en-CA" for demonstration
        // Установка локали на "en-CA" для демонстрации
        Api.SetLocale("en-CA");
        
        // Get the current locale
        // Получение текущей локали
        let locale = Api.GetLocale();
        
        // Display the locale in cell A1
        // Отображение локали в ячейке A1
        worksheet.GetRange("A1").SetValue("Locale: " + locale);
        
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
