#!/usr/bin/env python3
"""
servidor.py
Servidor HTTP local per al portafoli personal.

Executa'l des de la carpeta 'portafoli/' amb:
    python servidor.py

Després obre el navegador a:
    http://localhost:8000
"""

import http.server
import socketserver
import subprocess
import sys
import os
import webbrowser
from pathlib import Path

PORT = 8000
BASE = Path(__file__).parent


class PLMHandler(http.server.SimpleHTTPRequestHandler):
    """Servidor bàsic amb regeneració automàtica dels JSON abans de cada petició."""

    def do_GET(self):
        # Si es demana un data.json, regenerar-lo primer
        if self.path.endswith("data.json"):
            self._regenerar_json()
        super().do_GET()

    def _regenerar_json(self):
        script = BASE / "actualitzar.py"
        if script.exists():
            subprocess.run(
                [sys.executable, str(script)],
                capture_output=True,
                cwd=str(BASE)
            )

    def log_message(self, format, *args):
        # Silenciar els logs de cada petició (massa soroll)
        path = args[0] if args else ""
        if "data.json" in str(path) or ".html" in str(path):
            print(f"  → {args[0]}")


def main():
    os.chdir(BASE)

    print("\n" + "="*50)
    print("  Portafoli Personal — Servidor local")
    print("="*50)
    print(f"\n  Carpeta servida: {BASE}")
    print(f"  Adreça:          http://localhost:{PORT}")
    print(f"\n  Els JSON s'actualitzen automàticament")
    print(f"  cada cop que el navegador carrega una pàgina.")
    print(f"\n  Per aturar el servidor: Ctrl + C")
    print("="*50 + "\n")

    # Regenerar tots els JSON a l'inici
    script = BASE / "actualitzar.py"
    if script.exists():
        print("  Generant dades inicials...")
        result = subprocess.run(
            [sys.executable, str(script)],
            capture_output=True, text=True,
            cwd=str(BASE)
        )
        if result.stdout:
            for line in result.stdout.strip().split("\n"):
                if line.strip():
                    print(f"  {line}")
        print()

    # Obrir el navegador automàticament
    webbrowser.open(f"http://localhost:{PORT}")

    # Iniciar servidor
    with socketserver.TCPServer(("", PORT), PLMHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n\n  Servidor aturat. Fins aviat!\n")


if __name__ == "__main__":
    main()
