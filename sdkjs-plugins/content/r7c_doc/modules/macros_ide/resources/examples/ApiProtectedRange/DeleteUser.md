/**
 * @file DeleteUser_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiProtectedRange.DeleteUser
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to delete a user from a protected range.
 * It adds a protected range to the worksheet, adds a user to it, and then deletes that user.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как удалить пользователя из защищенного диапазона.
 * Он добавляет защищенный диапазон на лист, добавляет в него пользователя, а затем удаляет этого пользователя.
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
        // This example deletes the the user protected range.
        
        // How to close an access for the protected range to user specifing user id, name and access type.
        
        // Get an active sheet, add protected range to it, add users with rights then delete one of them. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.AddProtectedRange("protectedRange", "$A$1:$B$1");
        let protectedRange = worksheet.GetProtectedRange("protectedRange");
        protectedRange.AddUser("userId", "name", "CanView");
        protectedRange.DeleteUser("userId");
        
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
