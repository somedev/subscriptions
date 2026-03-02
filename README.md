# 💳 Семейный трекер подписок

Система для учёта семейных подписок и регулярных платежей на базе Google Sheets + Apps Script + Google Calendar.

## Возможности

- Список подписок с категориями, суммами, периодами и датами оплаты
- Подтверждение оплаты через чекбокс — автоматически логирует платёж и переносит дату
- Синхронизация с Google Calendar — события с push-уведомлениями
- Email-напоминания за N дней до оплаты
- История всех платежей
- Статистика: суммы по категориям и членам семьи
- Боковая панель для удобного добавления подписок
- Условное форматирование: просроченные (красный), скоро (оранжевый/жёлтый)

## Стек

- **Google Sheets** — хранение данных
- **Google Apps Script** (ES6, V8) — автоматизация
- **Google Calendar API** — push-уведомления
- **HTML Service** — форма добавления подписок

## Структура

```
src/
├── Code.gs              # onOpen(), кастомное меню
├── Config.gs            # Константы: листы, колонки, справочники
├── Setup.gs             # initialSetup() — создание листов, форматирование, триггеры
├── DailyCheck.gs        # dailyCheck() — ежедневная проверка и напоминания
├── CalendarSync.gs      # syncCalendar() — синхронизация с Google Calendar
├── PaymentConfirm.gs    # confirmPayment() — подтверждение оплаты
├── Notifications.gs     # sendReminder() — email-уведомления
├── Statistics.gs        # updateStatistics() — обновление дашборда
├── Utils.gs             # Вспомогательные функции
├── AddSubscription.gs   # Боковая панель добавления подписки
└── AddSubscriptionForm.html
appsscript.json          # Манифест: scopes, timezone, V8 runtime
```

## Деплой

Требуется [clasp](https://github.com/google/clasp) и Node.js.

```bash
# Установить clasp
npm install -g @google/clasp

# Войти в Google аккаунт
clasp login

# Включить Apps Script API: https://script.google.com/home/usersettings

# Создать новый проект (создаст .clasp.json)
clasp create --type sheets --title "💳 Подписки" --rootDir src

# Загрузить код
clasp push --force
```

После этого:

1. Открыть таблицу в Google Drive
2. Запустить `initialSetup()` через **Extensions > Apps Script**
3. Разрешить все запрошенные права
4. Перезагрузить таблицу — появится меню **"💳 Подписки"**
5. Заполнить **Настройки**: email(ы) и члены семьи

## Первый запуск

После `initialSetup()`:
1. Добавить подписку через **"➕ Добавить подписку"** в меню
2. Запустить **"📅 Синхронизировать календарь"**
3. Поделиться таблицей и календарём с членами семьи
