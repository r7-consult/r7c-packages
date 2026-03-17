---
interactive:
  version: 1
  enabled: true
  lessonId: "test-macros-base"
  title: "Тест: Базовые макросы"
  draft: false
  test:
    timeLimitSeconds: 120
    maxAttempts: 3
    passingScore: 60
    questions:
      - id: "q1"
        question: "Какая функция используется для получения активного листа?"
        options:
          - id: "a"
            text: "Api.GetSheet()"
          - id: "b"
            text: "Api.GetActiveSheet()"
          - id: "c"
            text: "Api.ActiveSheet()"
        correctOption: "b"
        topicRef: "Методы объекта Api (GetActiveSheet)"
      
      - id: "q2"
        question: "С помощью какого метода можно записать значение в ячейку?"
        options:
          - id: "a"
            text: "cell.SetValue()"
          - id: "b"
            text: "cell.PutValue()"
          - id: "c"
            text: "cell.Write()"
        correctOption: "a"
        topicRef: "Запись и чтение значений (SetValue)"
      
      - id: "q3"
        question: "Как правильнее всего перебрать все ячейки в диапазоне?"
        options:
          - id: "a"
            text: "Использовать цикл for...in"
          - id: "b"
            text: "Метод range.ForEach()"
          - id: "c"
            text: "Использовать while"
        correctOption: "b"
        topicRef: "Циклы и перебор элементов (ForEach)"

  task:
    goal: "Напишите макрос, закрепляющий знания теста."
    steps:
      - "Получите активный лист"
      - "Запишите в A1 текст 'Тест пройден'"
    acceptance:
      - "Ячейка A1 содержит текст 'Тест пройден'"
      
  starterCode: |
    (function() {
      // Ваш код здесь
    })();

  checks:
    - id: "check_a1_value"
      type: "cell_value"
      sheet: "Sheet1"
      cell: "A1"
      expected: "Тест пройден"
---

# Итоговый тест: Базовые макросы
Этот урок состоит из тестирования на время и практического задания для закрепления. Удачи!
