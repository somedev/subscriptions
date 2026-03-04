# CLAUDE.md — Family Subscription Tracker

## What is this project?

A Google Sheets + Apps Script + Google Calendar system for tracking family subscriptions and recurring payments. All UI and content is in **Russian**.

## Tech Stack

- **Google Sheets** — data storage, formulas, conditional formatting
- **Google Apps Script** (JavaScript ES6-compatible, V8 runtime) — automation, triggers, calendar sync, email notifications
- **Google Calendar API** (via Apps Script CalendarApp) — push notifications on all devices
- **HTML Service** (Apps Script) — sidebar forms for data entry
- **Web App** (Apps Script doGet) — mobile-friendly SPA interface

## Project Structure

```
src/
├── Code.gs                  # Entry point: onOpen(), doGet(), custom menu
├── Config.gs                # Constants: sheet names, column indices, defaults
├── Setup.gs                 # initialSetup() — creates sheets, formatting, triggers
├── DailyCheck.gs            # dailyCheck() — daily trigger, reminder logic
├── CalendarSync.gs          # syncCalendar() — calendar CRUD
├── PaymentConfirm.gs        # confirmPayment() — mark paid, log, advance date
├── Notifications.gs         # sendReminder() — email notifications
├── Statistics.gs             # updateStatistics() — dedicated statistics sheet
├── Utils.gs                 # Helpers: addMonths(), getSettings(), etc.
├── AddSubscription.gs       # Sidebar dialog for adding subscriptions
├── AddSubscriptionForm.html # HTML form for the sidebar
├── WebApi.gs                # Server-side API for the mobile web app
├── Index.html               # Web app shell: HTML structure, nav, views
├── Stylesheet.html          # Mobile-first CSS (included via template)
└── JavaScript.html          # Client-side routing, rendering, API calls
appsscript.json              # Manifest with scopes, timezone, webapp config
```

## Key Design Decisions

1. **All text in Russian** — sheet headers, menu items, email templates, calendar events, error messages, toast notifications — everything user-facing is in Russian
2. **Checkbox-driven workflow** — user checks "Оплачено" checkbox → onEdit trigger fires → payment logged to history → next date calculated → checkbox unchecked automatically
3. **Calendar as notification layer** — we create a dedicated calendar "💳 Подписки" and add all-day events with popup reminders so users get native push notifications on iPhone, Mac, and web
4. **No external dependencies** — pure Apps Script, no npm packages, no external APIs
5. **Script Properties for state** — use `PropertiesService.getScriptProperties()` to track last reminder dates and avoid duplicate notifications
6. **Mobile Web App** — deployed as Apps Script Web App via `doGet()`, vanilla JS SPA with onclick-based routing, card-based UI, bottom navigation. Single `getAllDataApi()` call fetches all data at once, client-side caching for instant tab switching

## Important Constants (Config.gs)

```javascript
const SHEET_SUBSCRIPTIONS = 'Подписки';
const SHEET_HISTORY = 'История оплат';
const SHEET_SETTINGS = 'Настройки';
const SHEET_LOOKUPS = 'Справочники';
const SHEET_STATISTICS = '📊 Статистика';

// Column indices (0-based) for "Подписки" sheet
const COL = {
  ID: 0,              // A
  NAME: 1,            // B - Название
  CATEGORY: 2,        // C - Категория
  AMOUNT: 3,          // D - Сумма
  CURRENCY: 4,        // E - Валюта
  PERIOD: 5,          // F - Период
  NEXT_DATE: 6,       // G - Дата следующей оплаты
  STATUS: 7,          // H - Статус
  LAST_PAID: 8,       // I - Последняя оплата
  IS_PAID: 9,         // J - Оплачено (checkbox)
  NOTIFICATIONS: 10,  // K - Уведомления
  REMIND_DAYS: 11,    // L - Дней до напоминания
  PAYER: 12,          // M - Кто платит
  PAY_METHOD: 13,     // N - Способ оплаты
  NOTES: 14,          // O - Примечания
  MONTHLY_COST: 15,   // P - Сумма/мес (formula)
  DAYS_UNTIL: 16,     // Q - Дней до оплаты (formula)
  CALENDAR_ID: 17     // R - Calendar Event ID (hidden)
};
```

## Common Patterns

### Getting a sheet
```javascript
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName(SHEET_SUBSCRIPTIONS);
```

### Reading subscription data
```javascript
const data = sheet.getDataRange().getValues();
const headers = data[0];
// data[1] onward = subscription rows
for (let i = 1; i < data.length; i++) {
  const row = data[i];
  if (!row[COL.NAME]) continue; // skip empty rows
  const name = row[COL.NAME];
  const amount = row[COL.AMOUNT];
  // ...
}
```

### Date arithmetic (month-safe)
```javascript
function addMonths(date, months) {
  const result = new Date(date);
  const day = result.getDate();
  result.setMonth(result.getMonth() + months);
  // Handle month-end overflow (e.g., Jan 31 + 1 month)
  if (result.getDate() !== day) {
    result.setDate(0); // Go to last day of previous month
  }
  return result;
}
```

### Calendar event creation
```javascript
const calendar = CalendarApp.getCalendarsByName(calendarName)[0];
const event = calendar.createAllDayEvent(title, date);
event.setDescription(description);
event.removeAllReminders();
event.addPopupReminder(daysBeforeInMinutes); // e.g., 3 days = 4320 min
event.addPopupReminder(0); // on the day
```

## Web App Architecture

The mobile web app (`doGet()` → `Index.html`) uses a single-page architecture:

- **One server call** — `getAllDataApi()` returns subscriptions + stats + history + settings in a single `google.script.run` round-trip
- **Client-side cache** — data stored in JS `state` object; tab switches render from cache instantly
- **Direct `google.script.run` callbacks** — no Promises or `.apply()` (unreliable in Apps Script sandbox)
- **onclick routing** — no `hashchange` events (unreliable in sandbox iframe); nav tabs use `onclick` → `showView()`
- **Refresh after mutations** — `refreshData()` re-fetches all data after add/edit/delete/pay operations
- **Edit from cache** — tapping a subscription card fills the edit form from cached data (no extra server call)

### Web App API functions (WebApi.gs)

| Function | Purpose |
|----------|---------|
| `getAllDataApi()` | Returns all data in one call (used on initial load) |
| `getSubscriptionApi(rowIndex)` | Single subscription (fallback for edit) |
| `addSubscriptionApi(data)` | Add new subscription |
| `updateSubscriptionApi(rowIndex, data)` | Update existing subscription |
| `deleteSubscriptionApi(rowIndex)` | Delete subscription + calendar event |
| `markAsPaidApi(rowIndex)` | Confirm payment (reuses `confirmPayment()`) |

## Testing

- Run `initialSetup()` first — it creates everything
- Add test subscriptions manually or via sidebar
- Check "Оплачено" on a test row — verify history log and date advance
- Run `syncCalendar()` — verify calendar events created
- Run `dailyCheck()` manually — check email delivery
- Verify conditional formatting colors
- Deploy web app and test on mobile: all 4 views, payment confirmation, add/edit/delete

## Gotchas & Tips

- Apps Script uses **server-side JavaScript** — no DOM, no `window`, no `fetch`. Use `UrlFetchApp` for HTTP calls
- `onEdit()` simple trigger can't access services requiring authorization. Use an **installable** onEdit trigger instead
- `MailApp.sendEmail()` has a daily quota: 100 emails for free accounts, 1500 for Workspace
- Calendar popup reminders are specified in **minutes** (3 days = 3 * 24 * 60 = 4320 minutes)
- `CalendarApp.getCalendarsByName()` returns an array — always use `[0]`
- When hiding a column: `sheet.hideColumns(columnIndex)` — 1-based index
- Dates in Sheets come as JavaScript `Date` objects when read via `getValues()`
- For checkboxes, use `sheet.insertCheckboxes()` or data validation with TRUE/FALSE
- Always use `Utilities.formatDate(date, 'Europe/Minsk', 'dd.MM.yyyy')` for date formatting
- Formulas use **European locale**: `;` as argument separator, `,` as decimal separator (e.g., `=IF(A1>0;A1*4,33;0)`)
- Custom periods supported: besides standard (Месяц, Квартал, etc.), users can enter "N мес." (e.g., "2 мес.", "5 мес.")
- **Web App sandbox limitations**: `addMetaTag()` only supports `viewport` — Apple meta tags must go directly in HTML `<head>`. `hashchange` and `Promise.apply()` are unreliable — use onclick handlers and direct `google.script.run` calls instead
- **Web App scope**: requires `spreadsheets` (not `spreadsheets.currentonly`) in manifest for `doGet()` context. `getActiveSpreadsheet().getUrl()` may fail — wrap in try-catch

## Scope Boundaries

**In scope (v1):**
- CRUD for subscriptions (add, edit via sheet, delete by changing status)
- Manual payment confirmation via checkbox
- Payment history log
- Calendar sync with push reminders
- Email notifications
- Statistics dashboard
- Conditional formatting
- Russian localization
- Mobile web app (view, add, edit, delete, pay, stats, settings)

**Out of scope (v2+):**
- Telegram bot
- Charts/graphs
- Bank CSV import
- Budget threshold alerts
