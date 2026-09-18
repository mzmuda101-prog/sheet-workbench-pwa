#!/usr/bin/env python3
"""Serwer statyczny pod testy Playwrighta (domyślnie port 4175).

Dlaczego nie `python3 -m http.server`: tamten odpowiada w HTTP/1.0, czyli bez
keep-alive — zamyka połączenie po każdej odpowiedzi. Przy dwudziestu zasobach
strony i kilkunastu testach pod rząd Chromium regularnie trafiał na zamknięte
gniazdo i przewracał losowy test błędem `net::ERR_CONNECTION_RESET` (za każdym
razem inny, a w izolacji wszystkie przechodziły). HTTP/1.1 + wątki to naprawiają.
"""
import sys
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer


class Handler(SimpleHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def end_headers(self):
        # Testy muszą widzieć świeże pliki po każdej zmianie w kodzie.
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, *args):
        pass  # cisza — logi żądań zalewałyby wyjście testów


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 4175
    ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
