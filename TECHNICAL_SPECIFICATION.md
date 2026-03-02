# Technical Specification: Family Subscription & Recurring Payments Tracker

## Project Overview

A family subscription and recurring payment tracking system built on Google Sheets + Google Apps Script + Google Calendar. The system provides data entry, viewing, statistics, payment confirmation, and automated reminders — all accessible from Mac, iPhone, and Web.

**Language:** Russian (all UI labels, headers, notifications, calendar events)
**Target users:** Family members (2–5 people) sharing a single Google Sheet

---

## Architecture

```
┌─────────────────────────────────────────────────────┐
│                  Google Sheet                        │
│                                                      │
│  Sheet 1: "Подписки" (Subscriptions)                │
│  Sheet 2: "История оплат" (Payment History)         │
│  Sheet 3: "Настройки" (Settings)                    │
│  Sheet 4: "Справочники" (Dropdowns/Lookups)         │
│                                                      │
├─────────────────────────────────────────────────────┤
│                Google Apps Script                     │
│                                                      │
│  - Daily trigger: check upcoming payments            │
│  - Create/update Google Calendar events              │
│  - Send email reminders                              │
│  - Auto-calculate next payment dates                 │
│  - Custom menu for quick actions                     │
│                                                      │
├─────────────────────────────────────────────────────┤
│              Google Calendar                          │
│                                                      │
│  - Dedicated calendar: "💳 Подписки"                 │
│  - Events with reminders (push notifications)        │
│  - Shared with family members                        │
│                                                      │
└─────────────────────────────────────────────────────┘
```

---

## Data Model

### Sheet 1: "Подписки" (Main Subscriptions Table)

| Column | Header (RU) | Type | Description |
|--------|-------------|------|-------------|
| A | ID | Auto-number | Unique identifier (auto-generated, do not edit) |
| B | Название | Text | Subscription/payment name (e.g., "Netflix", "Аренда", "Spotify Family") |
| C | Категория | Dropdown | Category: Стриминг, Софт, Связь, Жильё, Страховка, Здоровье, Образование, Другое |
| D | Сумма | Number (currency) | Payment amount per period |
| E | Валюта | Dropdown | RUB, USD, EUR, other |
| F | Период | Dropdown | Месяц, Квартал, Полгода, Год, Неделя |
| G | Дата следующей оплаты | Date | Next payment due date |
| H | Статус | Dropdown | Активна / Приостановлена / Отменена |
| I | Последняя оплата | Date | Date when last marked as paid |
| J | Оплачено | Checkbox | Manual checkbox — user marks payment as completed |
| K | Уведомления | Checkbox | TRUE = send reminders, FALSE = muted |
| L | Дней до напоминания | Number | How many days before payment to send reminder (default: 3) |
| M | Кто платит | Dropdown | Family member name (from Settings sheet) |
| N | Способ оплаты | Dropdown | Карта, Автосписание, Перевод, Наличные |
| O | Примечания | Text | Free-form notes |
| P | Сумма/мес (расчёт) | Formula | Normalized monthly cost (auto-calculated based on amount + period) |
| Q | Дней до оплаты | Formula | =G{row} - TODAY() — days until next payment |
| R | Calendar Event ID | Text (hidden) | Google Calendar event ID for sync (hidden column, managed by script) |

### Sheet 2: "История оплат" (Payment History Log)

| Column | Header (RU) | Type | Description |
|--------|-------------|------|-------------|
| A | ID записи | Auto-number | Unique log entry ID |
| B | ID подписки | Number | Reference to subscription ID from Sheet 1 |
| C | Название | Text | Subscription name (copied for readability) |
| D | Дата оплаты | Date | Actual payment date |
| E | Сумма | Number | Amount paid |
| F | Валюта | Text | Currency |
| G | Кто оплатил | Text | Family member who confirmed payment |
| H | Примечание | Text | Optional note |

### Sheet 3: "Настройки" (Settings)

| Row | Setting (Column A) | Value (Column B) |
|-----|-------------------|------------------|
| 1 | Название календаря | 💳 Подписки |
| 2 | Email для уведомлений | comma-separated emails |
| 3 | Время ежедневной проверки | 09:00 |
| 4 | Члены семьи | comma-separated names (e.g., "Алексей, Мария, Дети") |
| 5 | Валюта по умолчанию | RUB |

### Sheet 4: "Справочники" (Dropdown Data)

Contains named lists for data validation dropdowns:
- Column A: Категории (categories list)
- Column B: Периоды (Месяц, Квартал, Полгода, Год, Неделя)
- Column C: Валюты (RUB, USD, EUR)
- Column D: Способы оплаты (Карта, Автосписание, Перевод, Наличные)
- Column E: Статусы (Активна, Приостановлена, Отменена)

---

## Google Apps Script — Functional Requirements

### 1. Custom Menu

On spreadsheet open (`onOpen`), add a custom menu "💳 Подписки" with items:

- **"✅ Отметить оплату"** — Processes the currently selected row: sets "Оплачено" = TRUE, logs entry to Payment History, calculates and sets next payment date, then clears the checkbox
- **"📅 Синхронизировать календарь"** — Full sync: creates/updates/deletes calendar events for all active subscriptions
- **"📊 Обновить статистику"** — Refreshes calculated fields and dashboard area
- **"➕ Добавить подписку"** — Opens an HTML sidebar/dialog with a form for easy data entry
- **"⚙️ Первоначальная настройка"** — Creates the calendar, sets up triggers, initializes sheets if missing

### 2. Daily Trigger (Time-driven, runs once per day)

Function: `dailyCheck()`

Logic:
1. Read all rows from "Подписки" sheet where Статус = "Активна"
2. For each subscription where Уведомления = TRUE:
   a. Calculate days until next payment: `nextPaymentDate - today`
   b. If days <= "Дней до напоминания" AND days >= 0 AND reminder not yet sent for this cycle:
      - Send email notification to all configured recipients
      - Store last reminder date in Script Properties to avoid duplicates
3. For each subscription where "Оплачено" checkbox is TRUE (in case onEdit didn't fire):
   a. Run `confirmPayment()` for that row (safety net)

### 3. Calendar Sync

Function: `syncCalendar()`

Logic:
1. Get or create calendar by name from Settings sheet
2. For each row in "Подписки":
   a. If Статус = "Активна":
      - If column R (Event ID) is empty → create new all-day event:
        - Title: `"💳 {Название} — {Сумма} {Валюта}"`
        - Date: value from "Дата следующей оплаты"
        - Description: formatted string with all subscription details
        - Reminders: popup at N days before (from "Дней до напоминания"), and popup on the day (0 minutes)
        - Save event ID to column R
      - If column R has an event ID → update existing event (title, date, description, reminders)
   b. If Статус = "Приостановлена" or "Отменена":
      - If column R has an event ID → delete the calendar event
      - Clear column R

### 4. Payment Confirmation

Function: `confirmPayment(rowIndex)`

Triggered by: onEdit (checkbox toggle) or custom menu

Logic:
1. Read subscription data from the specified row
2. Validate: ensure the row has valid data (name, amount, date exist)
3. Append a new row to "История оплат" sheet:
   - Auto-increment ID
   - Copy: subscription ID, name, today as payment date, amount, currency
   - Set "Кто оплатил" = current user email or active family member
4. Calculate next payment date from current "Дата следующей оплаты":
   - Неделя → add 7 days
   - Месяц → add 1 month (handle month-end: Jan 31 + 1 month = Feb 28)
   - Квартал → add 3 months
   - Полгода → add 6 months
   - Год → add 1 year
5. Update the row:
   - "Дата следующей оплаты" = new calculated date
   - "Последняя оплата" = today
   - "Оплачено" = FALSE (uncheck)
6. Update or recreate calendar event for the new date

### 5. Email Notifications

Function: `sendReminder(subscriptionData, daysUntil)`

Uses `MailApp.sendEmail()`. Email content in Russian:

```
Subject: 💳 Напоминание: {Название} — оплата через {дней} дн.

Body:
Здравствуйте!

Напоминаем о предстоящей оплате:

📌 Подписка: {Название}
💰 Сумма: {Сумма} {Валюта}
📅 Дата оплаты: {Дата следующей оплаты}
👤 Кто платит: {Кто платит}
💳 Способ: {Способ оплаты}

⏰ Осталось дней: {дней}

Перейти к таблице: {spreadsheet URL}
```

Recipients: all emails from Settings row "Email для уведомлений"

### 6. Statistics / Dashboard

Create a summary section on the "Подписки" sheet below the data table (starting from row 50 or a clearly separated area), OR create a dedicated "Дашборд" sheet with:

- **Общая сумма/мес** — `=SUMIF(H:H,"Активна",P:P)` — total monthly cost
- **Общая сумма/год** — monthly × 12
- **По категориям** — SUMIFS grouped by category, showing monthly cost per category
- **По членам семьи** — SUMIFS grouped by "Кто платит"
- **Ближайшие 5 оплат** — Top 5 soonest payments (SORT + FILTER on active subscriptions by "Дней до оплаты")
- **Оплачено за текущий месяц** — SUMIFS on "История оплат" for current month
- **Всего активных подписок** — COUNTIF on Статус = "Активна"

### 7. onEdit Trigger (Installable)

Function: `onEditTrigger(e)`

Logic:
- Check if edited cell is in column J ("Оплачено") on the "Подписки" sheet
- Check if new value is TRUE (checkbox was checked)
- If yes → call `confirmPayment(e.range.getRow())`
- Show a toast notification: "✅ Оплата {Название} подтверждена. Следующая: {новая дата}"

### 8. Initial Setup

Function: `initialSetup()`

Logic:
1. Create sheets if they don't exist: "Подписки", "История оплат", "Настройки", "Справочники"
2. Write headers to each sheet
3. Set up data validation (dropdowns) on "Подписки" columns C, E, F, H, M, N using ranges from "Справочники"
4. Populate "Справочники" with default values
5. Populate "Настройки" with default settings
6. Apply conditional formatting to "Подписки":
   - Q <= 0 → Red background (#F4C7C3), bold
   - Q <= 3 → Orange background (#FCE8B2)
   - Q <= 7 → Yellow background (#FFF2CC)
   - H = "Отменена" → Strikethrough, gray text
   - H = "Приостановлена" → Italic, light gray
   - J = TRUE → Green background (#D9EAD3) for the entire row
7. Create Google Calendar "💳 Подписки" (or get existing by name)
8. Create daily time-driven trigger for `dailyCheck()` (delete existing first to avoid duplicates)
9. Create installable onEdit trigger for `onEditTrigger()`
10. Freeze row 1 (headers) on "Подписки" sheet
11. Hide column R (Calendar Event ID)
12. Set column widths for readability
13. Show completion dialog: "✅ Настройка завершена! Календарь создан, триггеры установлены."

---

## Conditional Formatting Rules (on "Подписки" sheet)

Applied to the data range (row 2 downward):

| Priority | Condition | Style |
|----------|-----------|-------|
| 1 | Custom formula: `=$H2="Отменена"` | Strikethrough, text color #999999 |
| 2 | Custom formula: `=$H2="Приостановлена"` | Italic, text color #AAAAAA |
| 3 | Custom formula: `=$J2=TRUE` | Background #D9EAD3 (green) |
| 4 | Custom formula: `=AND($H2="Активна",$Q2<=0)` | Background #F4C7C3 (red), bold |
| 5 | Custom formula: `=AND($H2="Активна",$Q2<=3,$Q2>0)` | Background #FCE8B2 (orange) |
| 6 | Custom formula: `=AND($H2="Активна",$Q2<=7,$Q2>3)` | Background #FFF2CC (yellow) |

---

## Formulas (pre-populated in headers, auto-filled for new rows)

### Column P: Monthly cost normalization
```
=IF(H2="Активна", SWITCH(F2, "Месяц",D2, "Квартал",D2/3, "Полгода",D2/6, "Год",D2/12, "Неделя",D2*4.33, 0), 0)
```

### Column Q: Days until payment
```
=IF(AND(H2="Активна", G2<>""), G2-TODAY(), "")
```

---

## Apps Script File Structure

```
project/
├── src/
│   ├── Code.gs              # onOpen, custom menu creation, entry points
│   ├── Config.gs            # Sheet names, column indices, default values (all as constants)
│   ├── Setup.gs             # initialSetup(), sheet creation, formatting, triggers
│   ├── DailyCheck.gs        # dailyCheck() — main daily trigger function
│   ├── CalendarSync.gs      # syncCalendar(), createEvent(), updateEvent(), deleteEvent()
│   ├── PaymentConfirm.gs    # confirmPayment(), calculateNextDate()
│   ├── Notifications.gs     # sendReminder(), email template
│   ├── Statistics.gs        # updateStatistics() — refresh dashboard formulas
│   ├── Utils.gs             # addMonths(), getSheetByName(), getSettings(), etc.
│   └── AddSubscription.gs   # showAddDialog(), processAddForm() — sidebar HTML form
├── src/html/
│   └── AddSubscription.html # HTML sidebar form for adding new subscriptions
└── appsscript.json          # Manifest: timezone, OAuth scopes
```

### Required OAuth Scopes (appsscript.json)

```json
{
  "timeZone": "Europe/Moscow",
  "dependencies": {},
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/mail.send"
  ],
  "exceptionLogging": "STACKDRIVER"
}
```

---

## Deployment Instructions

1. Create a new Google Spreadsheet
2. Open Extensions → Apps Script
3. Delete the default `Code.gs` content
4. Create all `.gs` files from the file structure above
5. Create `AddSubscription.html` in the HTML files section
6. Update `appsscript.json` with the manifest above
7. Save all files
8. Run `initialSetup()` from the script editor (or from the custom menu after reload)
9. Authorize all requested permissions when prompted
10. Reload the spreadsheet — the custom menu "💳 Подписки" should appear
11. Go to Settings sheet and configure: emails, family members
12. Add your first subscriptions
13. Run "📅 Синхронизировать календарь" from the menu
14. Share the spreadsheet and calendar with family members

---

## Error Handling

- All script functions must be wrapped in try-catch blocks
- Errors should be logged with `console.error()` and shown to user via `SpreadsheetApp.getUi().alert()`
- Calendar API errors (event not found, calendar deleted) should be handled gracefully: clear the stored Event ID and recreate
- MailApp quota exceeded: log error, skip notification, try again next day

---

## Edge Cases

1. **Month-end dates:** Jan 31 + 1 month = Feb 28 (or 29 in leap years). Use a robust `addMonths()` helper
2. **Duplicate triggers:** `initialSetup()` must delete existing triggers before creating new ones
3. **Concurrent edits:** Multiple family members editing simultaneously — Apps Script handles locking internally
4. **Empty rows:** Skip rows where "Название" is empty
5. **Calendar deleted externally:** If calendar is deleted by user, recreate it on next sync
6. **Event ID stale:** If event ID references a deleted event, catch the error, clear ID, create new event

---

## Testing Checklist

- [ ] initialSetup() creates all sheets, headers, formatting, triggers
- [ ] Adding a subscription row populates formulas in P and Q
- [ ] Checking "Оплачено" logs to history, advances date, unchecks
- [ ] Calendar sync creates events with correct dates and reminders
- [ ] Daily trigger sends email for subscriptions due within reminder window
- [ ] Email contains correct Russian text and subscription details
- [ ] Cancelled subscription removes calendar event
- [ ] Statistics show correct monthly/yearly totals
- [ ] Conditional formatting highlights overdue (red), soon (orange/yellow)
- [ ] Works on mobile Google Sheets app (columns not too wide)
- [ ] Menu items all function correctly
