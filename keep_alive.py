# -*- coding: utf-8 -*-
"""
Flowcats Anti-Sleep Guard · Render Keep-Alive Daemon
====================================================
Mantiene despierta tu aplicación en Render (o cualquier plataforma cloud)
haciendo peticiones HTTP regulares al endpoint /health cada 10 minutos
para evitar que el servidor gratuito se duerma por inactividad.

Uso:
    python keep_alive.py [URL_OPCIONAL]
    Ejemplo: python keep_alive.py https://flowcats-web.onrender.com
"""

import sys
import time
import os
import urllib.request
import urllib.error
from datetime import datetime

DEFAULT_INTERVAL_SECONDS = 600  # 10 minutos (Render duerme a los 15 min)

def get_target_url() -> str:
    if len(sys.argv) > 1 and sys.argv[1].strip():
        url = sys.argv[1].strip()
    else:
        url = os.environ.get("RENDER_EXTERNAL_URL") or os.environ.get("KEEP_ALIVE_URL") or ""
        
    if not url:
        # Intentar leer del .env
        env_path = os.path.join(os.path.dirname(__file__), ".env")
        if os.path.exists(env_path):
            with open(env_path, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("RENDER_EXTERNAL_URL=") or line.startswith("KEEP_ALIVE_URL="):
                        url = line.split("=", 1)[1].strip().strip('"\'')
                        break
                        
    if not url:
        url = "https://flowcats-web.onrender.com"
        
    if not url.startswith("http://") and not url.startswith("https://"):
        url = "https://" + url
        
    return url.rstrip("/")

def ping(target_url: str) -> tuple[bool, int, float, str]:
    health_url = f"{target_url}/health"
    start = time.time()
    try:
        req = urllib.request.Request(
            health_url,
            headers={"User-Agent": "Flowcats-KeepAlive/2.0 (Anti-Sleep Guard)"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            elapsed = round((time.time() - start) * 1000, 1)
            return True, resp.status, elapsed, "OK"
    except urllib.error.HTTPError as e:
        elapsed = round((time.time() - start) * 1000, 1)
        return False, e.code, elapsed, str(e.reason)
    except Exception as exc:
        elapsed = round((time.time() - start) * 1000, 1)
        return False, 0, elapsed, str(exc)

def main():
    target_url = get_target_url()
    print("=" * 65)
    print("  🐈‍⬛ FLOWCATS · ANTI-SLEEP GUARD (Render Keep-Alive)")
    print(f"  Objetivo : {target_url}/health")
    print(f"  Intervalo: Cada {DEFAULT_INTERVAL_SECONDS // 60} minutos")
    print("  Presiona CTRL + C para detener el guardián.")
    print("=" * 65)

    ping_count = 0
    while True:
        ping_count += 1
        now_str = datetime.now().strftime("%Y-%m-%d %I:%M:%S %p")
        success, code, ms, msg = ping(target_url)
        
        if success:
            status_tag = f"[✓] PING #{ping_count} EXITOSO"
            print(f"[{now_str}] {status_tag} · HTTP {code} ({ms}ms) -> Servidor despierto y activo")
        else:
            status_tag = f"[!] PING #{ping_count} ALERTA"
            print(f"[{now_str}] {status_tag} · HTTP {code} ({ms}ms) -> {msg}")

        # Esperar 10 minutos mostrando cuenta regresiva en consola
        for remaining in range(DEFAULT_INTERVAL_SECONDS, 0, -30):
            mins, secs = divmod(remaining, 60)
            time.sleep(30)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[!] Guardián anti-sleep detenido por el usuario.")
