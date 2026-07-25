"""Serve the Post E dashboard and a cached Micron stock-price API."""

from __future__ import annotations

import argparse
import json
import math
import threading
import time
from functools import partial
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import yfinance as yf


ROOT = Path(__file__).resolve().parent
STOCK_SYMBOL = "MU"
STOCK_CACHE_SECONDS = 60
_stock_cache: dict[str, object] = {"data": None, "expires_at": 0.0}
_stock_lock = threading.Lock()


def _sample_prices(values: list[float], limit: int = 48) -> list[float]:
    if len(values) <= limit:
        return values
    indexes = {
        round(index * (len(values) - 1) / (limit - 1))
        for index in range(limit)
    }
    return [values[index] for index in sorted(indexes)]


def _finite_float(value: object) -> float:
    numeric = float(value)
    if not math.isfinite(numeric):
        raise ValueError("Stock value is not finite")
    return numeric


def _download_micron_stock() -> dict[str, object]:
    ticker = yf.Ticker(STOCK_SYMBOL)
    history = ticker.history(
        period="5d",
        interval="30m",
        auto_adjust=False,
        actions=False,
        prepost=False,
    )
    if history.empty or "Close" not in history:
        history = ticker.history(
            period="7d",
            interval="1d",
            auto_adjust=False,
            actions=False,
        )
    if history.empty or "Close" not in history:
        raise RuntimeError("yfinance returned no Micron price history")

    close_series = history["Close"].dropna()
    closes = [_finite_float(value) for value in close_series.tolist()]
    if not closes:
        raise RuntimeError("yfinance returned no valid Micron close prices")

    daily_closes = close_series.groupby(close_series.index.date).last()
    price = _finite_float(daily_closes.iloc[-1])
    previous_close = (
        _finite_float(daily_closes.iloc[-2])
        if len(daily_closes) > 1
        else price
    )
    change = price - previous_close
    change_percent = (change / previous_close * 100) if previous_close else 0.0

    return {
        "symbol": STOCK_SYMBOL,
        "currency": "USD",
        "price": round(price, 4),
        "previous_close": round(previous_close, 4),
        "change": round(change, 4),
        "change_percent": round(change_percent, 4),
        "closes": [round(value, 4) for value in _sample_prices(closes)],
        "as_of": close_series.index[-1].isoformat(),
        "source": "yfinance",
        "stale": False,
    }


def get_micron_stock() -> tuple[dict[str, object], bool]:
    now = time.monotonic()
    cached = _stock_cache["data"]
    if cached is not None and now < float(_stock_cache["expires_at"]):
        return dict(cached), True

    with _stock_lock:
        now = time.monotonic()
        cached = _stock_cache["data"]
        if cached is not None and now < float(_stock_cache["expires_at"]):
            return dict(cached), True
        try:
            data = _download_micron_stock()
        except Exception:
            if cached is None:
                raise
            stale_data = dict(cached)
            stale_data["stale"] = True
            return stale_data, True

        _stock_cache["data"] = data
        _stock_cache["expires_at"] = now + STOCK_CACHE_SECONDS
        return dict(data), False


class DashboardHandler(SimpleHTTPRequestHandler):
    server_version = "PostEDashboard/1.0"

    def _send_json(self, payload: dict[str, object], status: HTTPStatus) -> None:
        body = json.dumps(payload, allow_nan=False, separators=(",", ":")).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/health":
            self._send_json(
                {"status": "ok", "service": "Post E Dashboard"},
                HTTPStatus.OK,
            )
            return
        if path == "/api/micron-stock":
            try:
                data, cached = get_micron_stock()
                data["cached"] = cached
                self._send_json(data, HTTPStatus.OK)
            except Exception as error:
                self._send_json(
                    {
                        "error": "Unable to retrieve Micron stock data",
                        "detail": str(error),
                        "source": "yfinance",
                    },
                    HTTPStatus.SERVICE_UNAVAILABLE,
                )
            return
        super().do_GET()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    handler = partial(DashboardHandler, directory=str(ROOT))
    server = ThreadingHTTPServer((args.host, args.port), handler)
    print(f"Post E Dashboard: http://{args.host}:{args.port}/Post%20E%20Dashboard.html")
    print(f"Micron stock API: http://{args.host}:{args.port}/api/micron-stock")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
