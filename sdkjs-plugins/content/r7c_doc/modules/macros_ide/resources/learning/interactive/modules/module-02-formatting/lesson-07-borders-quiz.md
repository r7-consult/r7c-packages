---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-07"
  title: "Урок 7: Quiz. Изменение границ ячеек"
  draft: false
  quiz:
    question: "Мы написали макрос, чтобы окрасить фон ячейки C3 в зеленый цвет. Какой кусок кода содержит фатальную ошибку, из-за которой макрос выдаст 'undefined' или вообще упадет при выполнении?"
    options:
      - id: "a"
        text: "let color = Api.CreateColorFromRGB(0, 255, 0);"
      - id: "b"
        text: "let sheet = Api.GetActiveSheet();"
      - id: "c"
        text: "sheet.GetRange(\"C3\").SetFillColor(\"green\");"
      - id: "d"
        text: "sheet.GetRange(\"C3\").SetBold(true);"
    correctOption: "c"
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

# Урок 7. Quiz: Изменение границ и форматирования

В прошлом уроке мы подробно разбирали объекты цвета и настройки шрифтов. Пришло время проверить вашу внимательность!

**Что мы проверяем:**
Понимание того, какие типы данных принимают методы Р7 API. В программировании очень важна строгая типизация. Если метод ждет объект, а вы передаете ему текстовую строку — скрипт неизбежно остановится с ошибкой.

**Задание:**
Представьте, что перед вами код из четырех строк, призванный выделить ячейку **C3** зеленым жирным шрифтом:

```javascript
let color = Api.CreateColorFromRGB(0, 255, 0); // Строка 1
let sheet = Api.GetActiveSheet();              // Строка 2
sheet.GetRange("C3").SetFillColor("green");    // Строка 3
sheet.GetRange("C3").SetBold(true);            // Строка 4
```

Где-то здесь кроется грубая ошибка (опечатка), свойственная новичкам, пришедшим из веб-программирования (HTML/CSS). Выберите строчку, которая сломает этот макрос!

*(Пояснения откроются сразу после вашего ответа)*
