# Локальное разворачивание enterprise-каталога

Этот POC раздает `r7c-packages` как обычный статический HTTP-каталог. Так можно
проверить корпоративный сценарий, где рабочее место пользователя не ходит на
GitHub за списком и файлами плагинов.

## Запуск

```powershell
cd C:\Users\Даниил\Desktop\plugins\r7c-packages
powershell -ExecutionPolicy Bypass -File .\enterprise\serve-static-catalog.ps1 -Port 8090
```

Сервер раздает корень репозитория по адресу:

```text
http://127.0.0.1:8090/
```

## Проверка endpoint-ов

```powershell
Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8090/health
Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8090/store/config.json
Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8090/sdkjs-plugins/content/hello-world/config.json
```

## Что должно быть видно в логе

При открытии `{r7} consult` в R7 Desktop сервер должен печатать запросы:

```text
GET /store/config.json
GET /sdkjs-plugins/content/<plugin>/config.json
GET /sdkjs-plugins/content/<plugin>/resources/...
```

Если эти строки появляются, каталог и ресурсы плагинов идут с локального
сервера, а не с GitHub.

## Связанный билд менеджера

Для проверки использовать билд из соседнего репозитория:

```text
C:\Users\Даниил\Desktop\plugins\r7c\build\r7c_enterprise_server_v1.1.4_local.plugin
```

В этом билде уже прописан `catalogBaseUrl=http://127.0.0.1:8090/`.
