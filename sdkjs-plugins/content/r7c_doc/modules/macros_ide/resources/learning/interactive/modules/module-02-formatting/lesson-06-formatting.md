---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-06"
  title: "Урок 6: Форматирование. Эквиваленты Interior.Color и Font"
  draft: false
  beforeScript: |
    (function() {
    try {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("G1:J1").Merge(false);
        sheet.GetRange("G1:J1").SetValue("💡 Интерактивная песочница готова. Напишите макрос!");
        sheet.GetRange("G1:J1").SetFontColor(Api.CreateColorFromRGB(120, 120, 120));
        sheet.GetRange("G1:J1").SetItalic(true);
    } catch(e) {}
    })();
---

# Урок 6. Форматирование: Эквиваленты VBA Interior и Font

Добро пожаловать в Модуль 2. Здесь мы разберем объектную модель форматирования в API Р7-Офис и сравним её с привычными свойствами Microsoft Excel VBA.

## 🎯 Архитектура форматирования
В VBA вы привыкли обращаться к свойствам объекта напрямую: `Range("A1").Interior.Color` или `Range("A1").Font.Bold = True`.
JavaScript API Р7-Офис использует инкапсулированные методы-сеттеры (`Set...`) и геттеры (`Get...`) объекта `ApiRange`.

## 🎨 Создание цвета: `Api.CreateColorFromRGB()`
В отличие от VBA, где можно использовать функцию `RGB(r, g, b)` или константы `vbRed`, в Р7-Офис цветовые объекты должны быть явно инстанцированы через главный класс `Api`.

```javascript
// Инстанцирование RGB-объектов
let redColor = Api.CreateColorFromRGB(255, 0, 0);
let greenColor = Api.CreateColorFromRGB(138, 181, 155);
```

---

## 🖌️ Таблица миграции методов форматирования

| Свойство VBA | Метод Р7-Офис API | Описание |
|--------------|-------------------|----------|
| `Range.Interior.Color` | `.SetFillColor(colorObject)` | Заливка фона ячейки объектом цвета. |
| `Range.Font.Color` | `.SetFontColor(colorObject)` | Изменение цвета текста. |
| `Range.Font.Bold` | `.SetBold(boolean)` | Установка жирного начертания (`true` / `false`). |
| `Range.Font.Size` | `.SetFontSize(number)` | Установка размера шрифта (в пунктах). |

### Синтаксический пример из Эталона
```javascript
(function() {
    'use strict';
    try {
        const api = Api;
        if (!api) throw new Error('API not available');
        
        let sheet = Api.GetActiveSheet();
        let cell = sheet.GetRange("A1");
        
        cell.SetValue("WARNING");
        
        let bgRed = Api.CreateColorFromRGB(255, 0, 0);       
        let textWhite = Api.CreateColorFromRGB(255, 255, 255); 
        
        // Применение методов форматирования в Ядре
        cell.SetFillColor(bgRed);         
        cell.SetFontColor(textWhite);     
        cell.SetBold(true);               
        cell.SetFontSize(16);             
        
    } catch(e) { /* Обработчик ошибок опущен для краткости */ }
})();
```

## ⚠️ Ловушки при миграции
1. **Строковые константы:** Метод `.SetFillColor()` не принимает HEX-строки (`"#FF0000"`) или имена (`"red"`). Только объект, созданный через `CreateColorFromRGB`.
2. **Аргументы методов:** В VBA `Font.Bold` — это свойство (Property), принимающее присваивание `= True`. В JS `.SetBold()` — это функция (Method), требующая передачи аргумента: `SetBold(true)`. Обязательно передавайте boolean-значение.

Переходите к следующему уроку, где мы закрепим работу с визуальной моделью данных.
