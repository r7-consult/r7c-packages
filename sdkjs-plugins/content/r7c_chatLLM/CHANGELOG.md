# Changelog

Все изменения проекта `{r7c}.ChatLLM` задокументированы в этом файле.
Проект следует [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

- Нет незарелизенных изменений.

---

## [2.1.0] - 2026-04-06

Релиз `2.1.0` фиксирует изменения из диапазона `b2998256716712e8c9e9fef0a2098f46e3b0db46..5843bea7567088553444a67640e5e562d6a6c189`.
Это релиз про надежность агентных сценариев, управляемый контекст, читаемый execution trace и переход spreadsheet-автоматизации на tool-first модель.

### Главное для пользователя

- Spreadsheet Agent теперь в первую очередь использует встроенные команды и host tools, а не свободно сгенерированные JS-макросы. Это снижает число ложных шагов и делает работу с таблицами более предсказуемой.
- Execution Trace переработан из сырого debug-лога в продуктовый pipeline: появились summary bar, понятные стадии, карточки шагов, более аккуратные состояния `stopped` и `error`, а также связка trace с итоговым ответом модели.
- Для Word добавлен отдельный внутренний snapshot tool, который снимает структурированное состояние документа и подает агенту надежный контекст вместо разовых макросов на чтение текста.
- Появился approval-first flow для Word: план сначала показывается пользователю, а подтверждение переводится в явную команду на исполнение вместо немедленного авто-запуска.
- В trace добавлен показ краткого reasoning summary модели. Для OpenRouter и OpenAI появились настройки `reasoningEffort`, а показ reasoning можно включать и выключать отдельным флагом.
- Улучшена работа с длинным контекстом: агент теперь умеет компактно сжимать историю, переживать token overflow и сохранять полезный контекст без развала длинных сессий.
- Улучшена работа с вложенными workbook-файлами: добавлен XLSX parser, предпросмотр, импорт данных из вложений и специализированный быстрый путь переноса данных в активную таблицу.

### Что изменилось по мотивам рабочих чатов

- Вопрос о том, почему модель не "помнит" прошлые ответы без передачи истории, привел к усилению контекстного менеджмента и переходу на встроенный сбор контекста `collect_context` вместо слабого prompt-time injection.
- Просьба сделать русский язык ответов дефолтным для основной аудитории из России и СНГ отражена в system prompt и planner policy: по умолчанию модель отвечает по-русски, но уважает явный запрос на другой язык.
- Просьба показать, как модель "думает", реализована через отдельные trace-записи `model_reasoning`, настройки показа reasoning и настройку `reasoningEffort` для поддерживаемых провайдеров.
- Поиск reasoning-capable моделей в OpenRouter сам по себе код не менял, но релиз подготовил инфраструктуру для тестирования таких моделей внутри плагина.

### Архитектурные и продуктовые изменения

- Добавлены новые модули `scripts/agent/r7chat_fastpath.js`, `scripts/platform/r7chat_desktop_tools_bridge.js`, `scripts/platform/r7chat_host_tools_client.js`, `scripts/shared/r7chat_xlsx_parser.js` и `scripts/ui/r7chat_trace_presenter.js`.
- Встроенный сбор контекста стал базовым механизмом для spreadsheet- и Word-сценариев: агенту теперь проще читать реальные данные редактора через доверенный путь, а не через произвольные макросы.
- Настройки UI упрощены и перестроены вокруг реальных пользовательских сценариев: reasoning toggle, reasoning effort, перегруппированные provider fields, обновленный inline settings panel и AI-first data context panel.
- Существенно расширен набор тестов для `agent_runtime`, `context_runtime`, `trace_presenter`, `openrouter_client`, `web_tools_client`, `desktop_tools_bridge` и настроек.
- В релиз попали сопроводительные дизайн-спеки и implementation plans по trace redesign, Word snapshot tool, desktop tools bridge и tool-first migration.

### Коммитный журнал релиза

- `2d2805a` `Implement built-in context flow and JS macro fast paths`: заменен старый prompt-time context injection на встроенный `collect_context`, добавлены bounded document/sheet collection, fast paths для типовых JS-макросов, bridge для desktop tools, обновлены настройки и UI панели контекста. Объем: `20` файлов, `+4528/-138`.
- `c762d27` `Implement Word plan approval flow and harden macro execution`: добавлен approval-first режим для Word-планов, карточки плана и image slots, подтверждение плана переведено в явную команду, усилена защита macro generation и recovery policy, обновлен composer UX и тесты. Объем: `7` файлов, `+1373/-23`.
- `a2b072d` `docs: Refactor CHANGELOG.md to follow semantic versioning`: changelog приведен к semver-формату и подготовлен к ведению релизами вместо произвольных заметок. Объем: `1` файл, `+66/-89`.
- `765e769` `Add model reasoning traces and context management`: реализованы reasoning traces, настройки `trace.showModelReasoning` и `providers.<id>.reasoningEffort`, контекстное сжатие истории и recovery после token overflow, обновлена языковая политика на default Russian для RU/CIS, усилена работа web tools. Объем: `16` файлов, `+1187/-49`.
- `e64fb81` `fix(ui): refine trace state mapper and add active step slide animation`: доработан mapper состояний trace, улучшено отображение активного шага и обновлена анимация pipeline. Объем: `2` файла, `+190/-123`.
- `4ab6bf5` `chore: ignore local worktrees directory`: добавлен `.gitignore` для локальной рабочей директории worktrees, чтобы технические файлы не шумели в git. Объем: `1` файл, `+1/-0`.
- `3c5d2e8` `Add robust Word document snapshot tool`: добавлен отдельный step type для Word snapshot, чтобы агент читал документ через надежный внутренний инструмент с метаданными, excerpt-ами и признаками структуры. Объем: `6` файлов, `+521/-22`.
- `6cee68e` `feat(ui): redesign execution trace and harden spreadsheet context collection`: execution trace полностью переработан в новый UI pipeline, добавлен `trace presenter`, улучшены summary/labels/details, ужесточен trusted path для spreadsheet context collection, обновлен settings panel и связка trace с ответом AI. Объем: `12` файлов, `+2200/-624`.
- `5843bea` `Implement tool-first spreadsheet automation`: spreadsheet-agent переведен на tool-first архитектуру, добавлен runtime host tool client, XLSX parser, path для импорта вложенных workbook-файлов, recovery policies для чтения/записи диапазонов, обновлен trace и сопроводительная документация. Объем: `34` файла, `+6338/-190`.

### Итог по масштабу релиза

- Всего в релиз вошло `9` коммитов после базового коммита `b299825`.
- Суммарно изменено `56` файлов.
- Суммарный дифф релиза: `+16275/-1129`.
- Самый большой объем работ пришелся на `r7chat_agent_runtime`, `r7chat_context_runtime`, `r7chat_ui`, `style.css`, `r7chat_planner_profiles` и новые supporting modules вокруг trace, host tools и XLSX parsing.

---

## [2.0.0] - 2026-04-04

### 🚀 Полные возможности плагина (Что умеет плагин)

#### 🤖 Мульти-провайдер AI
- **8 провайдеров:** OpenRouter, OpenAI, Anthropic, Gemini, YandexGPT, GigaChat, Mistral, DeepSeek.
- **On-demand discovery:** Автоматическая загрузка списка моделей от провайдера.
- **Режимы:** Streaming (обычный чат) и Non-streaming (tool loop).

#### 📊 Spreadsheet AI & Agent
- **Формулы:** `R7_ASK`, `R7_TRANSLATE`, `R7_EXTRACT`, `R7_CLASSIFY`, `R7_SUMMARIZE`.
- **Agent Runtime:** Многошаговое выполнение с трассировкой, повтором шагов и рестартом.
- **Контекст:** Инъекция контекста книги/листа в запросы.

#### 🌐 Web Tools (Поиск и Краулинг)
- **Провайдеры:** Exa (семантический поиск), Brave (веб-поиск).
- **Инструменты:** `web_search`, `web_crawling` доступны модели в чате.
- **Безопасность:** Блокировка localhost, loopback и частных IP.

#### 💬 Чат и Треды
- **Треды:** Множество независимых чатов, переключение, поиск.
- **Управление:** Авто-именование, ручное переименование, удаление.
- **Экспорт:** Выгрузка истории в `.docx`.

#### 📎 Вложения и Vision
- **Файлы:** DOCX, XLSX, PPTX, PDF.
- **Недавние:** Список последних 8 файлов.
- **Vision:** Анализ изображений (PNG, JPEG, WebP) через вставку или загрузку.

#### 🎨 UI/UX
- **Язык:** Переключатель (EN / RU / Как в редакторе).
- **Тема:** Темная / Светлая (авто + ручной выбор).
- **Контекстное меню:** Правка, объяснение, саммари, перевод, генерация картинок.
- **Slider AI:** Быстрый доступ к порталу.
