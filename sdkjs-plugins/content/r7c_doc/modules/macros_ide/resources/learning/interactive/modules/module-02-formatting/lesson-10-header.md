---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-10"
  title: "Урок 10: Практика. Создание стилизованной шапки отчета (Header Generation)"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        try {
            const api = Api;
            if (!api) throw new Error('API not available');
            
            let sheet = Api.GetActiveSheet();
            
            // ВАША ЗАДАЧА НИЖЕ:
            // 1. Получите объект ApiRange для "A1:E1"
            
            // 2. Выполните слияние ячеек (аналог Range.Merge)
            
            // 3. Заполните объединенный Range текстом "ГОДОВОЙ ОТЧЕТ"
            
            // 4. Установите свойства Font: Bold и Size (16)
            
            // 5. Выровняйте текст (Horizontal и Vertical Center)
            
            // 6. Установите заливку фона (FillColor) и текста (FontColor) 
            let black = Api.CreateColorFromRGB(0, 0, 0);
            let white = Api.CreateColorFromRGB(255, 255, 255);
            
        } catch (error) {
            console.error(error.message);
        }
    })();
  checks:
    - type: "range_values"
      range: "A1"
      expected: "ГОДОВОЙ ОТЧЕТ" 
  beforeScript: |
    (function() {
    let sheet = Api.GetActiveSheet();
    let headers = ["ID", "Имя", "Отдел", "Зарплата", "Статус"];
    for(let c=0; c<headers.length; c++) {
        sheet.GetRangeByNumber(0, c).SetValue(headers[c]);
    }
    sheet.GetRange("A4:F4").Merge(false);
    sheet.GetRange("A4:F4").SetValue("💡 Сырые данные в первой строке. Превратите их в красивую шапку!");
    })();
---

# Урок 10. Практика Инженерии: Программная верстка отчетов (UI Generation)

Объединение (Merge), заливка, шрифты — это стандартные методы `ApiRange`. При генерации отчетов "на лету" (Dynamic Reporting) скрипт зачастую формирует заголовочные структуры (Headers) до того, как вываливает матрицу данных.

В этом уроке вы закрепите макрос-вёрстку. Вы напишете кусок Enterprise-скрипта, генерирующий "Шапку" для таблицы.

## 📝 Требования (ТЗ на Форматирование)
Справа находится редактор с заготовкой IIFE. Выполните инстанцирование первой строки и стилизуйте её:

1. Инстанцируйте диапазон `A1:E1`. Рекомендуется присвоить его переменной (локальный кэш), чтобы избежать избыточных вызовов и длинных цепочек: `let header = sheet.GetRange("A1:E1")`.
2. Вызовите метод слияния `.Merge(false)`. Аргумент `false` (по умолчанию) означает слияние всего блока в одну гигантскую ячейку. Значение `true` сольет ячейки только построчно (Range.MergeAcross).
3. Присвойте `SetValue("ГОДОВОЙ ОТЧЕТ")`.
4. Измените начертание: `.SetBold(true)` и размер `.SetFontSize(16)`.
5. Установите центрирование по обеим осям. Аналоги VBA `HorizontalAlignment = xlCenter`: 
   - `.SetAlignHorizontal("center")`
   - `.SetAlignVertical("center")`
6. Примените инстанцированные цвета (из переменных `black` и `white`) для заливки `.SetFillColor()` и шрифта `.SetFontColor()`.

## ✍️ Эталон (Справочник методов)

```javascript
(function() {
    'use strict';
    try {
        const api = Api;
        let sheet = Api.GetActiveSheet();
        let header = sheet.GetRange("A1:E1"); // Кэшируем объект Range
        
        header.Merge(false);
        header.SetValue("ГОДОВОЙ ОТЧЕТ");
        
        // UI Форматирование (Font Properties)
        header.SetBold(true);
        header.SetFontSize(16);
        
        // UI Выравнивание (Alignment)
        header.SetAlignHorizontal("center");
        header.SetAlignVertical("center");
        
        // UI Заливка (Interior/Color)
        header.SetFillColor(Api.CreateColorFromRGB(0, 0, 0));
        header.SetFontColor(Api.CreateColorFromRGB(255, 255, 255));
    } catch(e) {}
})();
```

Выполните код в редакторе (Run), чтобы Р7-Офис отрендерил этот участок по вашим правилам и автотесты загорелись зеленым. Далее переходим к итоговому тесту Модуля 2.
