# Post E Dashboard

A glass-style operations dashboard for weekly Post E tracking, QDR insights,
issue trends, response-time analytics, reminders, feedback evidence, on-hold
status, and Micron stock monitoring.

## Run locally

The live Micron (`MU`) stock card uses a small Python API powered by
`yfinance`.

```powershell
py -m pip install -r requirements.txt
py server.py
```

Open:

```text
http://127.0.0.1:8765/Post%20E%20Dashboard.html
```

Opening the HTML file directly still loads the dashboard, but the live stock
card requires `server.py`.

## Run for other PCs on the same network

On the host PC:

```powershell
py server.py --host 0.0.0.0 --port 8765
```

Other PCs can open:

```text
http://HOST-PC-IP:8765/Post%20E%20Dashboard.html
```

Only the host PC needs Python, the packages in `requirements.txt`, and internet
access for `yfinance`.

## Main files

- `Post E Dashboard.html` — primary dashboard
- `server.py` — local web server and cached Micron stock API
- `requirements.txt` — Python dependency
- `Post E Dashboard Data Template*.xlsx` — dashboard data templates
- `Assets/` — dashboard imagery and logos
