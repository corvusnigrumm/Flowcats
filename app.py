# -*- coding: utf-8 -*-
"""
app.py — Servidor Web para Render y Plataformas Cloud.
Ejecuta exactamente los mismos métodos que la aplicación de escritorio Flowcats:
- Búsqueda RSS de El Tiempo y Portafolio.
- Titulares Groq AI de 3 palabras clave sin stop words.
- Temas del Día (Top 8 noticias con titulares <= 25 caracteres en TEMAS DEL DÍA.xlsx).
- Descarga directa de archivos Excel desde el navegador.
"""

import os
import sys
import io
import json
import threading
import time
from typing import List
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Importar funciones compartidas del motor de automatización
from automatizacion_santamaria import (
    run_scraper_selected,
    get_groq_api_key,
    export_temas_del_dia
)

app = FastAPI(
    title="Flowcats Web - Generador de Noticias",
    description="Plataforma Web para extracción de Flowcards y Temas del Día con IA Groq.",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Estado global de la tarea en ejecución
state_lock = threading.Lock()
execution_state = {
    "running": False,
    "progress": 0,
    "status": "Esperando inicio...",
    "logs": [],
    "last_updated": time.time(),
    "files_generated": []
}

class LogRedirector:
    def write(self, message):
        if message and str(message).strip():
            msg_str = str(message)
            with state_lock:
                execution_state["logs"].append(msg_str)
                if len(execution_state["logs"]) > 200:
                    execution_state["logs"] = execution_state["logs"][-200:]
                
                import re
                match = re.search(r'\(([0-9]+)%\)', msg_str)
                if match:
                    try:
                        execution_state["progress"] = int(match.group(1))
                        execution_state["status"] = f"Procesando... ({match.group(1)}%)"
                    except Exception:
                        pass
                execution_state["last_updated"] = time.time()

    def flush(self):
        pass

class RunRequest(BaseModel):
    selected_sources: List[str] = ["El Tiempo", "Portafolio"]
    include_amp: bool = True

def background_scraper_task(selected_sources: List[str], include_amp: bool):
    global execution_state
    redirector = LogRedirector()
    original_stdout = sys.stdout
    
    with state_lock:
        execution_state["running"] = True
        execution_state["progress"] = 0
        execution_state["status"] = "Iniciando extracción RSS..."
        execution_state["logs"] = ["=== INICIANDO PROCESO EN RENDER WEB ===\n"]
        execution_state["files_generated"] = []

    sys.stdout = redirector
    try:
        run_scraper_selected(selected_sources, include_amp=include_amp)
        with state_lock:
            execution_state["progress"] = 100
            execution_state["status"] = "Proceso completado exitosamente"
            files = []
            if "El Tiempo" in selected_sources and os.path.exists("El Tiempo.xlsx"):
                files.append("El Tiempo.xlsx")
            if "Portafolio" in selected_sources and os.path.exists("Portafolio.xlsx"):
                files.append("Portafolio.xlsx")
            if os.path.exists("TEMAS DEL DÍA.xlsx"):
                files.append("TEMAS DEL DÍA.xlsx")
            execution_state["files_generated"] = files
    except Exception as exc:
        with state_lock:
            execution_state["status"] = f"Error: {exc}"
            execution_state["logs"].append(f"\n[ERROR CRÍTICO] {exc}\n")
    finally:
        sys.stdout = original_stdout
        with state_lock:
            execution_state["running"] = False
            execution_state["last_updated"] = time.time()

@app.post("/api/run")
def api_run_scraper(req: RunRequest, background_tasks: BackgroundTasks):
    with state_lock:
        if execution_state["running"]:
            raise HTTPException(status_code=400, detail="Ya hay un proceso en ejecución.")
    
    background_tasks.add_task(background_scraper_task, req.selected_sources, req.include_amp)
    return {"message": "Proceso iniciado en segundo plano", "sources": req.selected_sources}

@app.get("/api/status")
def api_get_status():
    with state_lock:
        has_groq = bool(get_groq_api_key())
        return {
            "running": execution_state["running"],
            "progress": execution_state["progress"],
            "status": execution_state["status"],
            "logs": execution_state["logs"],
            "groq_active": has_groq,
            "files": execution_state["files_generated"]
        }

@app.get("/api/download/{filename}")
def api_download_file(filename: str):
    filename_clean = os.path.basename(filename)
    allowed_files = ["El Tiempo.xlsx", "Portafolio.xlsx", "TEMAS DEL DÍA.xlsx"]
    if filename_clean not in allowed_files:
        raise HTTPException(status_code=400, detail="Archivo no permitido")
    
    file_path = os.path.abspath(filename_clean)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail=f"El archivo {filename_clean} no existe.")
    
    return FileResponse(
        path=file_path,
        filename=filename_clean,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/", response_class=HTMLResponse)
def root_page():
    html_content = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Flowcats Web — Generador de Noticias El Tiempo & Portafolio</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-primary: #0a0e17;
            --bg-card: #131b2e;
            --bg-input: #1b2640;
            --accent-green: #00fa9a;
            --accent-blue: #00b4d8;
            --text-main: #f1f5f9;
            --text-sub: #94a3b8;
            --border-color: #2a3859;
            --shadow-glow: 0 0 25px rgba(0, 250, 154, 0.15);
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }

        body {
            font-family: 'Outfit', -apple-system, BlinkMacSystemFont, sans-serif;
            background-color: var(--bg-primary);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 2rem 1rem;
        }

        .container {
            width: 100%;
            max-width: 900px;
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 20px;
            padding: 2.5rem;
            box-shadow: 0 20px 50px rgba(0,0,0,0.5), var(--shadow-glow);
            backdrop-filter: blur(10px);
        }

        header {
            display: flex;
            align-items: center;
            gap: 1.5rem;
            margin-bottom: 2rem;
            padding-bottom: 1.5rem;
            border-bottom: 1px solid var(--border-color);
        }

        .logo-icon {
            font-size: 3.5rem;
            background: linear-gradient(135deg, #00fa9a, #00b4d8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-text h1 {
            font-size: 2.2rem;
            font-weight: 700;
            letter-spacing: -0.5px;
            background: linear-gradient(90deg, #ffffff, #94a3b8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .header-text p {
            color: var(--text-sub);
            font-size: 0.95rem;
            margin-top: 0.2rem;
        }

        .badge-status {
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
            background: rgba(0, 250, 154, 0.1);
            color: var(--accent-green);
            padding: 0.4rem 0.8rem;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
            border: 1px solid rgba(0, 250, 154, 0.3);
            margin-top: 0.5rem;
        }

        .pulse-dot {
            width: 8px;
            height: 8px;
            background-color: var(--accent-green);
            border-radius: 50%;
            box-shadow: 0 0 10px var(--accent-green);
            animation: pulse 1.5s infinite;
        }

        @keyframes pulse {
            0% { transform: scale(0.95); opacity: 0.8; }
            50% { transform: scale(1.3); opacity: 1; }
            100% { transform: scale(0.95); opacity: 0.8; }
        }

        .controls-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            background: var(--bg-input);
            padding: 1.5rem;
            border-radius: 14px;
            border: 1px solid var(--border-color);
            margin-bottom: 1.5rem;
        }

        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 0.75rem;
            cursor: pointer;
        }

        .checkbox-group input[type="checkbox"] {
            width: 18px;
            height: 18px;
            accent-color: var(--accent-green);
            cursor: pointer;
        }

        .checkbox-group label {
            font-size: 1rem;
            color: var(--text-main);
            cursor: pointer;
            font-weight: 500;
        }

        .btn-primary {
            width: 100%;
            padding: 1rem;
            background: linear-gradient(135deg, #00fa9a, #00b4d8);
            color: #0a0e17;
            font-size: 1.1rem;
            font-weight: 700;
            border: none;
            border-radius: 12px;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: 0 4px 15px rgba(0, 250, 154, 0.3);
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 0.5rem;
        }

        .btn-primary:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(0, 250, 154, 0.5);
        }

        .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; }

        .progress-section { margin-top: 1.5rem; }

        .progress-bar-bg {
            width: 100%;
            height: 10px;
            background: var(--bg-input);
            border-radius: 10px;
            overflow: hidden;
            border: 1px solid var(--border-color);
            margin-bottom: 0.5rem;
        }

        .progress-bar-fill {
            height: 100%;
            width: 0%;
            background: linear-gradient(90deg, #00fa9a, #00b4d8);
            transition: width 0.4s ease;
            box-shadow: 0 0 12px var(--accent-green);
        }

        .status-text {
            font-size: 0.9rem;
            color: var(--text-sub);
            display: flex;
            justify-content: space-between;
            font-weight: 500;
        }

        .console-box {
            margin-top: 1.5rem;
            background: #05080f;
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 1rem;
            height: 240px;
            overflow-y: auto;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.85rem;
            color: var(--accent-green);
            line-height: 1.5;
            white-space: pre-wrap;
        }

        .downloads-container {
            margin-top: 1.5rem;
            display: flex;
            flex-wrap: wrap;
            gap: 0.75rem;
        }

        .btn-download {
            padding: 0.75rem 1.2rem;
            background: rgba(0, 180, 216, 0.15);
            color: var(--accent-blue);
            border: 1px solid rgba(0, 180, 216, 0.4);
            border-radius: 10px;
            text-decoration: none;
            font-weight: 600;
            font-size: 0.9rem;
            transition: all 0.2s ease;
            display: inline-flex;
            align-items: center;
            gap: 0.5rem;
        }

        .btn-download:hover {
            background: rgba(0, 180, 216, 0.3);
            color: #ffffff;
            transform: translateY(-2px);
        }

        .btn-download-gold {
            background: rgba(255, 215, 0, 0.15);
            color: #ffd700;
            border-color: rgba(255, 215, 0, 0.4);
        }

        .btn-download-gold:hover {
            background: rgba(255, 215, 0, 0.3);
            color: #ffffff;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <div class="logo-icon">🐈‍⬛</div>
            <div class="header-text">
                <h1>Flowcats Web</h1>
                <p>Generador de noticias El Tiempo & Portafolio (Groq AI 3 Palabras & Temas del Día)</p>
                <div class="badge-status">
                    <div class="pulse-dot"></div>
                    <span id="ai-status">Groq AI Activo</span>
                </div>
            </div>
        </header>

        <div class="controls-grid">
            <div class="checkbox-group">
                <input type="checkbox" id="check-eltiempo" checked>
                <label for="check-eltiempo">El Tiempo</label>
            </div>
            <div class="checkbox-group">
                <input type="checkbox" id="check-portafolio" checked>
                <label for="check-portafolio">Portafolio</label>
            </div>
            <div class="checkbox-group">
                <input type="checkbox" id="check-amp" checked>
                <label for="check-amp">Incluir URL AMP</label>
            </div>
        </div>

        <button id="btn-start" class="btn-primary" onclick="startProcess()">
            <span>🚀 Iniciar Búsqueda y Generación RSS</span>
        </button>

        <div class="progress-section">
            <div class="progress-bar-bg">
                <div id="progress-fill" class="progress-bar-fill"></div>
            </div>
            <div class="status-text">
                <span id="status-label">Estado: Esperando inicio...</span>
                <span id="pct-label">0%</span>
            </div>
        </div>

        <div id="downloads" class="downloads-container"></div>

        <div id="console" class="console-box">Espere un momento... Los logs de ejecución aparecerán aquí.</div>
    </div>

    <script>
        async function updateStatus() {
            try {
                const res = await fetch('/api/status');
                const data = await res.json();

                document.getElementById('status-label').innerText = 'Estado: ' + data.status;
                document.getElementById('pct-label').innerText = data.progress + '%';
                document.getElementById('progress-fill').style.width = data.progress + '%';

                const consoleBox = document.getElementById('console');
                consoleBox.innerText = data.logs.join('');
                consoleBox.scrollTop = consoleBox.scrollHeight;

                const btn = document.getElementById('btn-start');
                if (data.running) {
                    btn.disabled = true;
                    btn.innerHTML = '<span>⏳ Procesando Noticias...</span>';
                } else {
                    btn.disabled = false;
                    btn.innerHTML = '<span>🚀 Iniciar Búsqueda y Generación RSS</span>';
                }

                if (data.files && data.files.length > 0) {
                    const downloadsDiv = document.getElementById('downloads');
                    downloadsDiv.innerHTML = '';
                    data.files.forEach(f => {
                        const a = document.createElement('a');
                        a.href = '/api/download/' + encodeURIComponent(f);
                        a.className = f.includes('TEMAS') ? 'btn-download btn-download-gold' : 'btn-download';
                        a.innerHTML = '📥 Descargar ' + f;
                        downloadsDiv.appendChild(a);
                    });
                }

                if (data.running) {
                    setTimeout(updateStatus, 2000);
                }
            } catch (err) {
                console.error('Error obteniendo estado:', err);
            }
        }

        async function startProcess() {
            const et = document.getElementById('check-eltiempo').checked;
            const p = document.getElementById('check-portafolio').checked;
            const amp = document.getElementById('check-amp').checked;

            const selected = [];
            if (et) selected.push('El Tiempo');
            if (p) selected.push('Portafolio');

            if (selected.length === 0) {
                alert('Por favor selecciona al menos un medio (El Tiempo o Portafolio).');
                return;
            }

            document.getElementById('btn-start').disabled = true;
            document.getElementById('downloads').innerHTML = '';

            try {
                const res = await fetch('/api/run', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        selected_sources: selected,
                        include_amp: amp
                    })
                });

                if (!res.ok) {
                    const err = await res.json();
                    alert('Error: ' + (err.detail || 'No se pudo iniciar'));
                    return;
                }

                setTimeout(updateStatus, 1000);
            } catch (err) {
                alert('Error conectando con el servidor: ' + err);
            }
        }

        updateStatus();
    </script>
</body>
</html>
"""
    return HTMLResponse(content=html_content)

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
