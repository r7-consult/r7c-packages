---
interactive:
  version: 1
  enabled: true
  lessonId: "hybrid-001-formula-quiz"
  title: "Гибридный тест: Запись формулы"
  targetEditor: "cell"
  difficulty: "easy"
  quiz:
    question: "Какая функция используется для получения активного рабочего листа в макросах Р7?"
    options:
      - id: "a"
        text: "Api.GetDocument()"
      - id: "b"
        text: "Api.GetActiveSheet()"
      - id: "c"
        text: "Api.GetSheet('Active')"
    correctOption: "b"
  task:
    goal: "Напишите макрос, который запишет формулу СУММ(A1:A5) в ячейку A6."
    steps:
      - "Сначала ответьте на теоретический вопрос."
      - "Затем получите активный лист."
      - "Запишите формулу `=СУММ(A1:A5)` в ячейку A6."
    acceptance:
      - "В ячейке A6 должна быть формула =СУММ(A1:A5)"
      - "Результат в А6 должен быть равен 150 (Сумма 10+20+30+40+50)"
  starterCode: |
    (function () {
      const ws = Api.GetActiveSheet();
      // Ваш код здесь
      
    })();
  beforeScript: |
    (function () {
      const ws = Api.GetActiveSheet();
      ws.GetRange("A1:C20").Clear();
      
      ws.GetRange("A1").SetValue("10");
      ws.GetRange("A2").SetValue("20");
      ws.GetRange("A3").SetValue("30");
      ws.GetRange("A4").SetValue("40");
      ws.GetRange("A5").SetValue("50");
      
      // UI Уведомление для пользователя
      let notice = ws.GetRange("C1:E1");
      notice.Merge(false);
      notice.SetValue("✅ База сгенерирована. Начинайте писать макрос!");
      notice.SetFontColor(Api.CreateColorFromRGB(0, 150, 0));
      notice.SetBold(true);
    })();
  checks:
    - id: "check-formula"
      type: "cell_formula"
      sheet: "Active"
      cell: "A6"
      expected: "=СУММ(A1:A5)"
    - id: "check-sum"
      type: "cell_value"
      sheet: "Active"
      cell: "A6"
      expected: 150
      normalize: "number"
  hints:
    - "Тест можно пройти только выбрав правильный вариант ответа."
    - "Для вставки формулы используйте метод SetValue() для ячейки."
  source:
    category: "Гибрид"
    examplePath: ""
  draft: false
---
# Теория и Практика

В этом уроке вы познакомитесь с гибридным фоматом обучения. 
Сначала ответьте на тестовый вопрос ниже. После успешного ответа вам откроется редактор кода для выполнения небольшого практического задания на закрепление.

Удачи!
