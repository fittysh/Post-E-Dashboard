# EDS Response Tracker

A production-ready, glass-style issue-management dashboard for monitoring EDS
response times, target compliance, active work, QDR performance, operational
analytics, data quality, and management reporting.

## Open the dashboard

Open `Post E Dashboard.html` in a modern browser. The dashboard is a standalone
HTML application and stores issue records and preferences in browser storage.

For local HTTP access, run:

```powershell
py server.py
```

Then open:

```text
http://127.0.0.1:8765/Post%20E%20Dashboard.html
```

## Main capabilities

- Responsive glass interface with light and dark themes
- Issue creation, editing, filtering, search, and target monitoring
- QDR performance, response-time analytics, and operational insights
- Interactive PDF summaries and management reports
- CSV import/export and JSON backup/restore
- Protected workspace-data reset
- Compact Inter and Roboto Mono typography
- Accessible navigation, keyboard controls, and reduced-motion support

## Main files

- `Post E Dashboard.html` — primary EDS Response Tracker dashboard
- `Assets/eds-mark.png` — EDS navigation and browser-tab logo
- `server.py` — optional local web server
- `Post E Dashboard Data Template*.xlsx` — import templates retained for existing users
- `Post E Onhold Details.html` — legacy detail view retained for older bookmarks
