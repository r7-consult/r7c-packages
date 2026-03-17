---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-format-11"
  title: "Тест: Итоги 2 модуля. Форматирование стобцов"
  draft: false
  test:
    timeLimit: 360
    passingScore: 80
    attempts: 3
    questions:
      - id: "q1_rgb"
        text: "С помощью какого метода в API Р7-Офис необходимо создавать цвет перед его применением к ячейке?"
        options:
          - id: "a"
            text: "Api.NewColor('red')"
          - id: "b"
            text: "Api.CreateColorFromRGB(R, G, B)"
          - id: "c"
            text: "sheet.GetColor()"
          - id: "d"
            text: "Цвет можно задать просто словом 'red' внутри SetFillColor"
        correctOption: "b"
        explanation: "API Р7-Офис использует жесткую типизацию. В метод SetFillColor нельзя передать текст. Сначала через Api нужно создать объект цвета по модели RGB: CreateColorFromRGB(255, 0, 0)."

      - id: "q2_bold"
        text: "Какой правильный синтаксис для установки жирного шрифта в диапазоне?"
        options:
          - id: "a"
            text: "range.SetBold(true);"
          - id: "b"
            text: "range.SetBold();"
          - id: "c"
            text: "range.Bold = true;"
          - id: "d"
            text: "Api.SetBold(range, true);"
        correctOption: "a"
        explanation: "Методы форматирования вызываются у самого диапазона (ApiRange). Метод SetBold требует передачи аргумента типа boolean (true для включения, false для отключения)."

      - id: "q3_merge"
        text: "Что произойдет с данными в ячейках B1, C1 и D1, если объединить диапазон A1:D1 с помощью метода .Merge(false)?"
        options:
          - id: "a"
            text: "Их текст сольется в одну длинную фразу (конкатенация)"
          - id: "b"
            text: "Они будут удалены. Сохранятся данные только верхней левой ячейки (A1)"
          - id: "c"
            text: "Метод выдаст ошибку, если в ячейках не пустой текст"
          - id: "d"
            text: "Их данные перенесутся на вторую строку (A2...)"
        correctOption: "b"
        explanation: "Как и при ручном объединении ячеек в интерфейсе, макрос удаляет все данные, кроме тех, что находятся в верхней левой (первой) ячейке объединяемого диапазона."

      - id: "q4_chaining"
        text: "Какая из этих записей (сокращение вызовов) правильная, если мы хотим окрасить ячейку C5 и вписать текст за один заход?"
        options:
          - id: "a"
            text: "sheet.SetValue(\"ОК\").GetRange(\"C5\").SetFillColor(color);"
          - id: "b"
            text: "sheet.GetRange(\"C5\").SetFillColor(color).SetValue(\"ОК\");"
          - id: "c"
            text: "Api.SetFillColor(color).sheet.GetRange(\"C5\");"
          - id: "d"
            text: "В Р7-Офис нельзя объединять методы в цепочку, нужно писать в 2 строки."
        correctOption: "b"
        explanation: "Методы API Р7-Офис для объекта ApiRange (диапазона) возвращают сам этот диапазон после выполнения операции. Это позволяет выстраивать их в цепочку: Получить ячейку -> Изменить Цвет -> Вставить текст (Chaining)."

      - id: "q5_align"
        text: "Какой строковый аргумент необходимо передать в метод SetAlignHorizontal, чтобы текст располагался строго посередине ячейки?"
        options:
          - id: "a"
            text: "\"middle\""
          - id: "b"
            text: "\"center\""
          - id: "c"
            text: "\"central\""
          - id: "d"
            text: "\"justify\""
        correctOption: "b"
        explanation: "Для центрирования по горизонтали и вертикали в API Р7-Офис (Web-синтаксис) используется ключевое слово 'center'. Опция 'middle' в данном API не применяется."
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

# Итоговый тест Модуля 2

Ура! Вы завершили второй модуль, посвященный визуальному форматированию таблиц с помощью макросов.

### Вы научились:
- Менять фон ячеек и текст (`SetFillColor`, `SetFontColor`).
- Настраивать жирность шрифта и его размер (`SetBold`, `SetFontSize`).
- Создавать объекты цветов через RGB (`Api.CreateColorFromRGB()`).
- Писать скрипты условного форматирования с логическим ветвлением (`if...else`).
- Программно объединять ячейки в блоки (`Merge`) и центрировать их содержимое по горизонтали и вертикали (`SetAlignHorizontal` и `SetAlignVertical`).

### 🏆 Правила тестирования:
В тесте **5 проверяющих вопросов**, покрывающих теорию 2 модуля.
Как и прежде, вопросы требуют понимания кода, типитизации (разницы между строкой и объектом) и поведения методов Р7-Офис.

- Лимит: 6 минут.
- Проходной балл: 80% (можно совершить одну ошибку).
- Три попытки.

Желаем удачи! Готовы? Жмите "Начать тест".
