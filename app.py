# -*- coding: utf-8 -*-
"""
Flowcats v2.0 · Servidor Web para Render y Plataformas Cloud / Local
Mesa de Redacción Automatizada:
- Procesos separados e independientes: Flowcards SEO (máx 3 palabras) y Temas del Día (Top 8 <= 25c).
- Selección flexible de medios: El Tiempo y/o Portafolio.
- IA Groq con modelo openai/gpt-oss-120b para titulares periodísticos de alto impacto.
- Módulo Anti-Sleep Guard integrado para Render Cloud.
- Descarga directa e individual de libros Excel.
"""

import os
import sys
import io
import json
import threading
import time
import urllib.request
import urllib.parse
from typing import List, Dict, Any, Optional
from fastapi import FastAPI, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DIST_DIR = os.path.join(BASE_DIR, "dist")
os.makedirs(DIST_DIR, exist_ok=True)
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from automatizacion_santamaria import (
    run_scraper_selected,
    get_groq_api_key,
    export_temas_del_dia
)

app = FastAPI(
    title="Flowcats v2.0 — Sala de Redacción Automatizada",
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

# Estado global de ejecución
state_lock = threading.Lock()
execution_state: Dict[str, Any] = {
    "running": False,
    "progress": 0,
    "status": "Sistema en espera · listo para ejecutar",
    "logs": [
        "[✓] Flowcats v2.0 inicializado.",
        "[*] Esperando instrucciones del operador..."
    ],
    "last_updated": time.time(),
    "files_generated": [],
    "topics": {
        "top": [],
        "El Tiempo": [],
        "Portafolio": []
    }
}

class LogRedirector:
    def __init__(self, original_stdout):
        self.original_stdout = original_stdout

    def write(self, message):
        if self.original_stdout:
            try:
                self.original_stdout.write(message)
                self.original_stdout.flush()
            except Exception:
                pass

        if message and str(message).strip():
            msg_str = str(message).rstrip()
            with state_lock:
                execution_state["logs"].append(msg_str)
                if len(execution_state["logs"]) > 600:
                    execution_state["logs"] = execution_state["logs"][-600:]
                
                import re
                match = re.search(r'\(([0-9]+)%\)', msg_str)
                if match:
                    try:
                        p = int(match.group(1))
                        execution_state["progress"] = p
                        if p < 100:
                            execution_state["status"] = f"Procesando noticias... ({p}%)"
                    except Exception:
                        pass
                
                if "INICIANDO EXTRACCIÓN PARA: El Tiempo" in msg_str or "INICIANDO EXTRACCION PARA: El Tiempo" in msg_str:
                    execution_state["status"] = "Extrayendo noticias de El Tiempo..."
                elif "INICIANDO EXTRACCIÓN PARA: Portafolio" in msg_str or "INICIANDO EXTRACCION PARA: Portafolio" in msg_str:
                    execution_state["status"] = "Extrayendo grupos económicos de Portafolio..."
                elif "PROCESANDO TEMAS DEL DÍA" in msg_str:
                    execution_state["status"] = "Generando Temas del Día (Top 8 <= 25c con Groq openai/gpt-oss-120b)..."
                elif "Archivo El Tiempo.xlsx generado" in msg_str:
                    execution_state["status"] = "Empaquetando El Tiempo.xlsx..."
                elif "Archivo Portafolio.xlsx generado" in msg_str:
                    execution_state["status"] = "Empaquetando Portafolio.xlsx..."
                elif "TEMAS DEL DÍA.xlsx" in msg_str and "Guardado" in msg_str:
                    execution_state["status"] = "Guardando TEMAS DEL DÍA.xlsx..."
                
                execution_state["last_updated"] = time.time()

    def flush(self):
        if self.original_stdout:
            try:
                self.original_stdout.flush()
            except Exception:
                pass


class RunRequest(BaseModel):
    selected_sources: List[str] = ["El Tiempo", "Portafolio"]
    process_type: str = "both"  # "flowcards", "temas_del_dia", "both"
    include_amp: bool = True


def format_article_topic(art: dict, source_name: str) -> dict:
    title = str(art.get("titulo_temas_25") or art.get("titulo_flowcard") or art.get("titulo_3") or art.get("titulo_raw") or "Sin título").strip()
    category = str(art.get("categoria") or "GENERAL").strip().upper()
    score = int(art.get("seo_score") or 85)
    score = max(50, min(99, score))
    
    keywords = art.get("keywords") or []
    if not keywords and art.get("titulo_flowcard"):
        keywords = [w.lower() for w in art["titulo_flowcard"].split() if len(w) > 2][:3]
    if not keywords and title:
        words = [w.lower() for w in title.split() if len(w) > 4][:3]
        keywords = words
        
    medio_src = art.get("medio") or art.get("source_name") or source_name
    return {
        "t": title,
        "c": category,
        "h": score,
        "k": keywords[:3],
        "s": medio_src
    }


def background_scraper_task(selected_sources: List[str], process_type: str, include_amp: bool):
    global execution_state
    original_stdout = sys.stdout
    redirector = LogRedirector(original_stdout)
    
    run_flowcards = process_type in ("flowcards", "both")
    run_temas = process_type in ("temas_del_dia", "both")
    run_el_tiempo = "El Tiempo" in selected_sources
    run_portafolio = "Portafolio" in selected_sources

    with state_lock:
        execution_state["running"] = True
        execution_state["progress"] = 2
        execution_state["status"] = "Inicializando motores de extracción..."
        execution_state["logs"] = [
            "=== INICIANDO PROCESO EN FLOWCATS WEB ===",
            f"[*] Modo de proceso: {process_type.upper()} ({'Flowcards (máx 3 palabras)' if run_flowcards else ''}{' + ' if run_flowcards and run_temas else ''}{'Temas del Día (<= 25 chars)' if run_temas else ''})",
            f"[*] Fuentes seleccionadas: {', '.join(selected_sources)}",
            f"[*] Motor IA: Groq (openai/gpt-oss-120b) · {'Conectado' if get_groq_api_key() else 'Modo heurístico'}",
            f"[*] Incluir URLs AMP: {'Sí' if include_amp else 'No'}"
        ]
        execution_state["files_generated"] = []
        execution_state["topics"] = {"top": [], "El Tiempo": [], "Portafolio": []}

    sys.stdout = redirector
    try:
        results = run_scraper_selected(selected_sources, process_type=process_type, include_amp=include_amp)
        
        topics_et = []
        topics_pf = []
        topics_top = []

        if isinstance(results, dict):
            raw_et = results.get("El Tiempo", [])
            raw_pf = results.get("Portafolio", [])
            raw_top = results.get("top", [])
            
            if run_el_tiempo and raw_et:
                topics_et = [format_article_topic(a, "El Tiempo") for a in raw_et]
            
            if run_portafolio and raw_pf:
                topics_pf = [format_article_topic(a, "Portafolio") for a in raw_pf]

            if run_temas and raw_top:
                topics_top = [format_article_topic(a, a.get("medio") or "El Tiempo") for a in raw_top]

        # Detectar archivos generados según el tipo de proceso y medio (en dist o raíz)
        files = []
        if run_flowcards:
            if run_el_tiempo and (os.path.exists(os.path.join(DIST_DIR, "El Tiempo.xlsx")) or os.path.exists(os.path.join(BASE_DIR, "El Tiempo.xlsx"))):
                files.append("El Tiempo.xlsx")
            if run_portafolio and (os.path.exists(os.path.join(DIST_DIR, "Portafolio.xlsx")) or os.path.exists(os.path.join(BASE_DIR, "Portafolio.xlsx"))):
                files.append("Portafolio.xlsx")

        # Bug 5 fix: detectar TEMAS DEL DÍA.xlsx sin importar qué medio fue seleccionado
        if run_temas:
            for temas_name in ["TEMAS DEL DÍA.xlsx", "TEMAS DEL DIA.xlsx"]:
                if os.path.exists(os.path.join(DIST_DIR, temas_name)) or os.path.exists(os.path.join(BASE_DIR, temas_name)):
                    files.append(temas_name)
                    break

        with state_lock:
            execution_state["progress"] = 100
            execution_state["status"] = "Proceso completado exitosamente"
            execution_state["files_generated"] = files
            execution_state["topics"] = {
                "top": topics_top,
                "El Tiempo": topics_et,
                "Portafolio": topics_pf
            }
            execution_state["logs"].append("=== PROCESO COMPLETADO EXITOSAMENTE ===")
            
    except Exception as exc:
        with state_lock:
            execution_state["status"] = f"Error: {exc}"
            execution_state["logs"].append(f"[ERROR CRÍTICO] {exc}")
    finally:
        sys.stdout = original_stdout
        with state_lock:
            execution_state["running"] = False
            execution_state["last_updated"] = time.time()


@app.post("/api/run")
def api_run_scraper(req: RunRequest, background_tasks: BackgroundTasks):
    if not req.selected_sources:
        raise HTTPException(status_code=400, detail="Debes seleccionar al menos una fuente de noticias (El Tiempo o Portafolio).")
    
    if req.process_type not in ("flowcards", "temas_del_dia", "both"):
        req.process_type = "both"

    with state_lock:
        if execution_state["running"]:
            raise HTTPException(status_code=400, detail="Ya hay un proceso de extracción en ejecución.")
        execution_state["running"] = True

    background_tasks.add_task(background_scraper_task, req.selected_sources, req.process_type, req.include_amp)
    return {
        "message": "Proceso iniciado en segundo plano",
        "sources": req.selected_sources,
        "process_type": req.process_type,
        "include_amp": req.include_amp
    }


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
            "files": execution_state["files_generated"],
            "topics": execution_state["topics"]
        }


@app.get("/api/download/{filename:path}")
def api_download_file(filename: str):
    filename_clean = urllib.parse.unquote(filename).strip()
    filename_clean = os.path.basename(filename_clean)
    
    allowed_files = [
        "El Tiempo.xlsx",
        "Portafolio.xlsx",
        "TEMAS DEL DÍA.xlsx",
        "TEMAS DEL DIA.xlsx",
        "TEMAS_DEL_DIA.xlsx"
    ]
    
    found_path = None
    for cand in allowed_files:
        if cand.lower() == filename_clean.lower():
            p_dist = os.path.join(DIST_DIR, cand)
            p_base = os.path.join(BASE_DIR, cand)
            if os.path.exists(p_dist):
                found_path = p_dist
                filename_clean = cand
                break
            elif os.path.exists(p_base):
                found_path = p_base
                filename_clean = cand
                break
    
    if not found_path:
        for p in [os.path.join(DIST_DIR, filename_clean), os.path.join(BASE_DIR, filename_clean)]:
            if os.path.exists(p) and filename_clean.endswith(".xlsx"):
                found_path = p
                break

    if not found_path or not os.path.exists(found_path):
        raise HTTPException(status_code=404, detail=f"El archivo '{filename_clean}' no fue encontrado en el servidor.")

    return FileResponse(
        path=found_path,
        filename=filename_clean,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/ping")
@app.get("/health")
@app.get("/api/health")
def health_check():
    return {
        "status": "ok",
        "app": "Flowcats Web",
        "version": "2.0.0",
        "keep_alive": "active",
        "timestamp": time.time()
    }


@app.get("/", response_class=HTMLResponse)
def root_page():
    return HTML_CONTENT


# Background self-pinging thread for Render
def _start_internal_keep_alive():
    external_url = os.environ.get("RENDER_EXTERNAL_URL") or os.environ.get("KEEP_ALIVE_URL")
    if not external_url:
        return
    
    def _worker():
        time.sleep(60)  # Esperar a que el servidor termine de arrancar
        target = f"{external_url.rstrip('/')}/health"
        print(f"[Anti-Sleep Guard] Iniciado worker de auto-ping hacia {target} cada 10 min.")
        while True:
            try:
                req = urllib.request.Request(target, headers={"User-Agent": "Flowcats-SelfPing/2.0"})
                with urllib.request.urlopen(req, timeout=20) as resp:
                    pass
            except Exception:
                pass
            time.sleep(600)  # Cada 10 minutos
            
    t = threading.Thread(target=_worker, daemon=True)
    t.start()

_start_internal_keep_alive()


HTML_CONTENT = r'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="theme-color" content="#0B0F19">
<meta name="description" content="Flowcats · Sala de redacción automatizada. Noticias y Flowcards SEO con FastAPI, Render y Groq AI.">
<title>Flowcats — Sala de Redacción Automatizada</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🐈‍⬛</text></svg>">

<!-- Tipografías: Fraunces (titular editorial) + Inter (UI) + Fira Code (datos) -->
<link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@fontsource/fraunces@5/index.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@fontsource/inter@5/index.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@fontsource/fira-code@5/index.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css">

<style>
/* ============================================================
   FLOWCATS v2.0 · "Mesa de redacción" — dark editorial console
   ============================================================ */
:root{
  --bg-0:#0B0F19;
  --bg-1:#111827;
  --panel:rgba(30,41,59,.66);
  --panel-deep:rgba(15,23,42,.55);
  --ink-deep:#070B14;
  --border:rgba(255,255,255,.08);
  --border-hi:rgba(255,255,255,.17);
  --emerald:#10B981;
  --spring:#00FA9A;
  --blue:#3B82F6;
  --violet:#8B5CF6;
  --amber:#FBBF24;
  --red:#F87171;
  --text:#E2E8F0;
  --muted:#94A3B8;
  --faint:#64748B;
  --dim:#475569;
  --paper:#F2EEE3;      /* blanco cálido de titulares */
  --serif:'Fraunces',Georgia,serif;
  --sans:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
  --mono:'Fira Code',Consolas,'Courier New',monospace;
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{
  font-family:var(--sans);
  background:linear-gradient(168deg,var(--bg-0) 0%,var(--bg-1) 100%);
  background-attachment:fixed;
  color:var(--text);
  min-height:100vh;overflow-x:hidden;
  -webkit-font-smoothing:antialiased;
}
::selection{background:rgba(16,185,129,.35)}
:focus-visible{outline:2px solid var(--spring);outline-offset:2px;border-radius:4px}
button{font-family:inherit}

/* ---------- Fondo: grano de imprenta + luz de mesa ---------- */
.bg{position:fixed;inset:0;z-index:-1;overflow:hidden;pointer-events:none}
.bg-noise{
  position:absolute;inset:0;opacity:.05;
  background-image:url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='160' height='160'><filter id='n'><feTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='2'/></filter><rect width='100%25' height='100%25' filter='url(%23n)' opacity='0.6'/></svg>");
}
.bg-lamp{
  position:absolute;top:-300px;left:50%;transform:translateX(-50%);
  width:1000px;height:640px;border-radius:50%;
  background:radial-gradient(closest-side,rgba(16,185,129,.13),transparent 72%);
}
.bg-rule-l,.bg-rule-r{position:absolute;top:0;bottom:0;width:1px;background:rgba(255,255,255,.035)}
.bg-rule-l{left:5%}.bg-rule-r{right:5%}

/* ============================================================
   CABECERA · mancha de periódico
   ============================================================ */
.masthead{
  max-width:1320px;margin:0 auto;
  padding:26px 28px 18px;
  display:flex;justify-content:space-between;align-items:flex-end;gap:18px;flex-wrap:wrap;
}
.mast-left{display:flex;align-items:center;gap:18px;min-width:0}
.cat-stamp{
  width:58px;height:58px;flex:none;display:grid;place-items:center;font-size:30px;
  border:1.5px solid rgba(242,238,227,.28);border-radius:6px;
  background:rgba(15,23,42,.5);
  transform:rotate(-3deg);
  box-shadow:3px 3px 0 rgba(0,0,0,.45),0 0 26px rgba(0,250,154,.12);
  transition:transform .35s cubic-bezier(.34,1.56,.64,1);
}
.cat-stamp:hover{transform:rotate(3deg) scale(1.05)}
.mast-title{display:flex;align-items:baseline;gap:12px;flex-wrap:wrap}
.mast-title h1{
  font-family:var(--serif);font-weight:900;
  font-size:clamp(30px,4.4vw,42px);letter-spacing:.5px;line-height:.95;
  color:var(--paper);
}
.ver-chip{
  font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:2px;
  color:#6EE7B7;border:1px solid rgba(16,185,129,.45);
  padding:3px 8px;border-radius:4px;background:rgba(16,185,129,.07);
  align-self:center;
}
.mast-sub{font-family:var(--mono);font-size:11px;letter-spacing:2.4px;text-transform:uppercase;color:var(--faint);margin-top:7px}
.mast-right{text-align:right}
.mast-date{font-family:var(--serif);font-style:italic;font-size:16px;color:var(--paper)}
.mast-edition{font-family:var(--mono);font-size:10px;letter-spacing:2.2px;color:var(--faint);margin-top:5px;text-transform:uppercase}
.double-rule{max-width:1320px;margin:0 auto;height:7px;border-top:3px solid rgba(255,255,255,.2);border-bottom:1px solid rgba(255,255,255,.2)}

/* ============================================================
   BARRA DE ESTADO + TICKER
   ============================================================ */
.statusbar{
  max-width:1320px;margin:0 auto;padding:12px 28px;
  display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;
}
.sb-left,.sb-right{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.sb-right{font-family:var(--mono);font-size:10.5px;letter-spacing:1.6px;color:var(--faint)}
.ai-badge{
  display:inline-flex;align-items:center;gap:8px;
  padding:7px 13px;border-radius:6px;font-size:12px;font-weight:600;
  border:1px solid;transition:all .4s ease;white-space:nowrap;
}
.ai-badge.groq{color:#6EE7B7;border-color:rgba(16,185,129,.45);background:rgba(16,185,129,.08)}
.ai-badge.heur{color:#FCD34D;border-color:rgba(251,191,36,.4);background:rgba(251,191,36,.06)}
.pulse-dot{width:8px;height:8px;border-radius:50%;background:var(--spring);flex:none;box-shadow:0 0 0 0 rgba(0,250,154,.6);animation:pulse 2s infinite}
.pulse-dot.amber{background:var(--amber);animation-name:pulseAmber}
@keyframes pulse{70%{box-shadow:0 0 0 7px rgba(0,250,154,0)}100%{box-shadow:0 0 0 0 rgba(0,250,154,0)}}
@keyframes pulseAmber{70%{box-shadow:0 0 0 7px rgba(251,191,36,0)}100%{box-shadow:0 0 0 0 rgba(251,191,36,0)}}
.server-pill{
  display:inline-flex;align-items:center;gap:8px;
  font-family:var(--mono);font-size:10.5px;letter-spacing:1.6px;color:var(--muted);
  padding:8px 12px;border-radius:6px;background:var(--panel-deep);border:1px solid var(--border);
}
.dot{width:7px;height:7px;border-radius:2px;background:var(--faint);flex:none}
.dot.green{background:var(--spring);box-shadow:0 0 9px var(--spring)}
.dot.amber{background:var(--amber);box-shadow:0 0 9px var(--amber)}
.rec{display:inline-flex;align-items:center;gap:6px;color:var(--red);font-weight:700}
.rec i{width:8px;height:8px;border-radius:50%;background:var(--red);animation:blinkRec 1s steps(2,start) infinite}
@keyframes blinkRec{to{visibility:hidden}}

.ticker{overflow:hidden;border-block:1px solid var(--border);background:rgba(11,15,25,.65)}
.ticker-track{display:flex;width:max-content;animation:marquee 48s linear infinite}
.ticker:hover .ticker-track{animation-play-state:paused}
@keyframes marquee{to{transform:translateX(-50%)}}
.tk-group{display:flex;align-items:center;gap:46px;padding:9px 23px 9px 46px}
.tk{display:inline-flex;align-items:center;gap:12px;white-space:nowrap;font-size:12.5px;color:#CBD5E1}
.tk b{font-family:var(--mono);font-size:9.5px;font-weight:700;letter-spacing:2px;color:var(--spring)}
.tk-dot{color:#334155;font-size:8px;margin-left:34px}

/* ============================================================
   CUBIERTA · grilla de paneles
   ============================================================ */
.deck{max-width:1320px;margin:0 auto;padding:24px 28px 8px}
.grid-main{display:grid;grid-template-columns:minmax(330px,400px) 1fr;gap:20px;align-items:start}
.stack{display:flex;flex-direction:column;gap:20px;min-width:0}
.grid-bottom{display:grid;grid-template-columns:minmax(300px,.85fr) 1.4fr;gap:20px;margin-top:20px;align-items:start}

.panel{
  position:relative;
  background:var(--panel);
  backdrop-filter:blur(16px);-webkit-backdrop-filter:blur(16px);
  border:1px solid var(--border);border-radius:10px;
  padding:20px 22px;
  box-shadow:0 12px 34px rgba(2,6,17,.45);
  opacity:0;transform:translateY(16px);
  transition:opacity .6s ease,transform .6s cubic-bezier(.22,1,.36,1),border-color .3s;
}
.panel.in{opacity:1;transform:none}
.panel:hover{border-color:var(--border-hi)}


.modes-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
.mode-card{
  display:flex;align-items:center;justify-content:space-between;gap:8px;
  padding:12px 13px;border-radius:8px;cursor:pointer;user-select:none;
  background:var(--panel-deep);border:1px solid var(--border);border-left:3px solid transparent;
  transition:transform .25s,border-color .25s,background .25s,box-shadow .25s;
  position:relative;
}
.mode-card:hover{transform:translateY(-2px);border-color:var(--border-hi);border-left-color:rgba(0,250,154,.4)}
.mode-card input{position:absolute;opacity:0;pointer-events:none}
.mode-card.selected{
  border-left-color:var(--spring);
  background:linear-gradient(90deg,rgba(16,185,129,.1),rgba(16,185,129,.02)),var(--panel-deep);
}
.mode-card.locked{opacity:.5;pointer-events:none}
.mode-card.selected .sc-check{
  background:var(--spring);border-color:transparent;color:#04120C;
  box-shadow:0 0 13px rgba(0,250,154,.5);transform:scale(1.06);
}
.sec-subhead{
  font-family:var(--mono);font-size:10px;letter-spacing:1.8px;
  color:var(--faint);text-transform:uppercase;margin-bottom:8px;
}

/* Etiqueta numerada de sección */
.panel-tag{
  display:flex;align-items:center;gap:10px;margin-bottom:16px;
  font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:2.4px;
  text-transform:uppercase;color:var(--muted);
}
.panel-tag .num{color:var(--spring);font-weight:700}
.panel-tag::after{content:'';flex:1;height:1px;background:linear-gradient(90deg,var(--border),transparent)}
.panel-head-row{display:flex;align-items:center;gap:12px;margin-bottom:16px;flex-wrap:wrap}
.panel-head-row .panel-tag{margin-bottom:0;flex:1;min-width:200px}

/* ============================================================
   01 · CONFIGURACIÓN
   ============================================================ */
.sources-grid{display:flex;flex-direction:column;gap:12px}
.source-card{
  display:flex;align-items:center;gap:16px;
  padding:15px 16px;border-radius:8px;cursor:pointer;user-select:none;
  background:var(--panel-deep);border:1px solid var(--border);border-left:3px solid transparent;
  transition:transform .25s,border-color .25s,background .25s,box-shadow .25s;
  position:relative;
}
.source-card:hover{transform:translateY(-2px);border-color:var(--border-hi);border-left-color:rgba(0,250,154,.4)}
.source-card input{position:absolute;opacity:0;pointer-events:none}
.source-card.selected{
  border-left-color:var(--spring);
  background:linear-gradient(90deg,rgba(16,185,129,.1),rgba(16,185,129,.02)),var(--panel-deep);
}
.source-card.locked{opacity:.5;pointer-events:none}
.source-card:has(input:focus-visible){outline:2px solid var(--spring);outline-offset:2px}
.sc-box{display:flex;flex-direction:column;gap:4px;min-width:0}
.sc-mast{line-height:1}
.sc-mast.et{font-family:var(--serif);font-weight:900;font-size:19px;letter-spacing:.5px;color:var(--paper)}
.sc-mast.pf{font-weight:800;font-size:16px;letter-spacing:3px;color:var(--paper)}
.sc-sub{font-family:var(--mono);font-size:9.5px;letter-spacing:1.8px;color:var(--faint);text-transform:uppercase}
.sc-check{
  margin-left:auto;width:21px;height:21px;border-radius:5px;flex:none;
  border:1.5px solid rgba(255,255,255,.22);
  display:grid;place-items:center;font-size:11px;color:transparent;
  transition:all .22s cubic-bezier(.34,1.56,.64,1);
}
.source-card.selected .sc-check{
  background:var(--spring);border-color:transparent;color:#04120C;
  box-shadow:0 0 13px rgba(0,250,154,.5);transform:scale(1.06);
}
.switch-row{
  display:flex;justify-content:space-between;align-items:center;gap:14px;
  margin-top:13px;padding:13px 15px;border-radius:8px;cursor:pointer;
  background:var(--panel-deep);border:1px solid var(--border);
  transition:border-color .25s,opacity .25s;
}
.switch-row:hover{border-color:rgba(59,130,246,.45)}
.switch-row.locked{opacity:.5;pointer-events:none}
.switch-info{display:flex;flex-direction:column;gap:2px}
.switch-info strong{font-size:13.5px;font-weight:600;display:flex;align-items:center;gap:7px}
.switch-info strong i{color:#60A5FA;font-size:12px}
.switch-info span{font-size:11px;color:var(--muted)}
.switch{position:relative;display:inline-flex;flex:none}
.switch input{position:absolute;opacity:0;width:100%;height:100%;cursor:pointer}
.track{
  width:46px;height:25px;border-radius:99px;display:block;position:relative;
  background:rgba(148,163,184,.2);border:1px solid var(--border);
  transition:all .3s ease;pointer-events:none;
}
.thumb{position:absolute;top:2px;left:2px;width:19px;height:19px;border-radius:50%;background:#94A3B8;transition:left .3s cubic-bezier(.34,1.56,.64,1),background .3s}
.switch input:checked + .track{background:linear-gradient(135deg,var(--emerald),var(--spring));border-color:transparent;box-shadow:0 0 14px rgba(16,185,129,.4)}
.switch input:checked + .track .thumb{left:23px;background:#fff}
.switch input:focus-visible + .track{outline:2px solid var(--spring);outline-offset:3px}

.btn-run{
  width:100%;margin-top:16px;padding:15px 22px;border:none;border-radius:8px;
  display:flex;align-items:center;justify-content:center;gap:10px;
  font-size:13.5px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#03130C;
  background:linear-gradient(120deg,var(--emerald),var(--spring) 38%,#34D399 66%,var(--emerald));
  background-size:230% 100%;background-position:0 0;
  cursor:pointer;position:relative;overflow:hidden;
  box-shadow:0 6px 24px rgba(16,185,129,.3),inset 0 1px 0 rgba(255,255,255,.35);
  transition:background-position .5s,transform .25s,box-shadow .25s,opacity .25s;
}
.btn-run:hover:not(:disabled){background-position:95% 0;transform:translateY(-2px);box-shadow:0 10px 34px rgba(0,250,154,.4),inset 0 1px 0 rgba(255,255,255,.4)}
.btn-run:active:not(:disabled){transform:translateY(0)}
.btn-run:disabled{opacity:.45;cursor:not-allowed;box-shadow:none}
.btn-run.loading{background:linear-gradient(120deg,#1E293B,#334155 50%,#1E293B);background-size:230% 100%;color:#94A3B8;animation:flow 2.4s linear infinite}
.run-hint{display:none;margin-top:11px;font-size:12px;color:var(--amber);align-items:center;gap:7px}
.run-hint.show{display:flex;animation:rowIn .3s ease}
.etiqueta{
  margin-top:13px;font-family:var(--serif);font-style:italic;font-size:12.5px;
  color:#8b8574;line-height:1.5;
}
.ripple{position:absolute;border-radius:50%;background:rgba(255,255,255,.45);transform:scale(0);animation:ripple .6s ease-out forwards;pointer-events:none}
@keyframes ripple{to{transform:scale(3.4);opacity:0}}
.spinner{width:16px;height:16px;flex:none;border-radius:50%;border:2px solid rgba(255,255,255,.25);border-top-color:currentColor;animation:spin .7s linear infinite;display:inline-block}
.spinner.sm{width:13px;height:13px;border-width:1.5px}
.spinner.big{width:28px;height:28px;border-width:3px;border-color:rgba(16,185,129,.18);border-top-color:var(--emerald);margin:0 auto 8px}
@keyframes spin{to{transform:rotate(360deg)}}

/* ============================================================
   02 · EJECUCIÓN
   ============================================================ */
.chip-time{
  margin-left:auto;display:inline-flex;align-items:center;gap:7px;
  font-family:var(--mono);font-size:11px;font-weight:600;letter-spacing:1.5px;color:var(--muted);
  padding:4px 10px;border-radius:6px;background:var(--panel-deep);border:1px solid var(--border);
}
.chip-time i{color:#34D399}
.progress-top{display:flex;justify-content:space-between;align-items:flex-end;gap:14px;margin-bottom:12px}
#actionText{font-size:13px;color:#CBD5E1;min-height:20px;line-height:1.5}
.pct-num{font-family:var(--mono);font-size:34px;font-weight:700;line-height:1;flex:none;color:var(--paper);font-variant-numeric:tabular-nums;text-shadow:0 0 22px rgba(0,250,154,.25)}
.bar{height:12px;border-radius:6px;overflow:hidden;position:relative;background:var(--ink-deep);border:1px solid rgba(255,255,255,.07)}
.fill{
  height:100%;width:0%;position:relative;
  background:repeating-linear-gradient(45deg,var(--emerald) 0 12px,var(--spring) 12px 24px);
  background-size:200% 100%;animation:stripes 1.1s linear infinite;
  box-shadow:0 0 16px rgba(16,185,129,.5);
  transition:width .55s cubic-bezier(.22,1,.36,1);
}
@keyframes stripes{to{background-position:34px 0}}
.metrics{display:flex;margin-top:18px;border-top:1px solid var(--border);border-bottom:1px solid var(--border)}
.metric{flex:1;padding:12px 8px 12px 14px;border-left:1px solid var(--border)}
.metric:first-child{border-left:none;padding-left:2px}
.metric .lbl{display:block;font-family:var(--mono);font-size:9px;letter-spacing:2.2px;color:var(--faint);text-transform:uppercase;margin-bottom:5px}
.metric b{font-family:var(--mono);font-size:19px;font-weight:700;font-variant-numeric:tabular-nums}
.metric b.small{font-size:12.5px;color:#6EE7B7;font-weight:600;letter-spacing:.5px}
.stepper{display:flex;align-items:center;gap:9px;margin-top:16px;flex-wrap:wrap}
.step{
  display:inline-flex;align-items:center;gap:7px;
  font-family:var(--mono);font-size:10px;letter-spacing:1.8px;color:var(--faint);
  padding:5px 9px;border:1px solid transparent;border-radius:6px;transition:all .4s ease;
}
.step i{font-size:9px}
.step.active{color:var(--spring);border-color:rgba(0,250,154,.4);text-shadow:0 0 12px rgba(0,250,154,.55)}
.step.done{color:#34D399}
.st-arrow{color:#334155;font-size:11px}

/* ============================================================
   03 · CONSOLA
   ============================================================ */
.p-terminal{padding-bottom:0;display:flex;flex-direction:column;overflow:hidden}
.p-terminal .panel-tag{padding:0 2px}
.term-head{
  display:flex;align-items:center;gap:12px;padding:11px 14px;flex-wrap:wrap;
  background:rgba(15,23,42,.85);border-top:1px solid var(--border);border-bottom:1px solid var(--border);
}
.dots{display:flex;gap:7px}
.dots span{width:12px;height:12px;border-radius:50%;transition:filter .2s}
.dots span:hover{filter:brightness(1.3)}
.dots .r{background:#FF5F57}.dots .y{background:#FEBC2E}.dots .g{background:#28C840}
.term-title{font-family:var(--mono);font-size:11.5px;color:var(--faint);display:flex;align-items:center;gap:8px;min-width:0}
.term-title i{color:#34D399}
.term-actions{margin-left:auto;display:flex;align-items:center;gap:8px}
.log-count{font-family:var(--mono);font-size:10px;color:var(--faint);letter-spacing:1px;margin-right:2px}
.term-btn{
  display:inline-flex;align-items:center;gap:6px;
  font-family:var(--mono);font-size:10px;font-weight:600;letter-spacing:1.2px;color:var(--muted);
  background:rgba(148,163,184,.07);border:1px solid var(--border);
  padding:6px 10px;border-radius:6px;cursor:pointer;transition:all .2s;
}
.term-btn:hover{color:#fff;border-color:rgba(16,185,129,.45);background:rgba(16,185,129,.12)}
.term-body{
  font-family:var(--mono);font-size:12.5px;line-height:1.8;
  padding:15px 18px;height:338px;overflow-y:auto;overscroll-behavior:contain;
  background:radial-gradient(120% 90% at 15% 0%,rgba(16,185,129,.05),transparent 42%),var(--ink-deep);
}
.term-body::-webkit-scrollbar{width:8px}
.term-body::-webkit-scrollbar-thumb{background:rgba(148,163,184,.2);border-radius:99px}
.term-body::-webkit-scrollbar-thumb:hover{background:rgba(16,185,129,.45)}
.term-line{display:flex;gap:12px;animation:rowIn .25s ease both}
@keyframes rowIn{from{opacity:0;transform:translateY(6px)}}
.log-time{color:#3F4C63;flex:none;font-size:11px;padding-top:1px}
.log-text{word-break:break-word;min-width:0;color:#CBD5E1}
.lg-head .log-text{color:#A78BFA;font-weight:700;letter-spacing:.6px}
.lg-error .log-text{color:var(--red);font-weight:600}
.lg-warn .log-text{color:var(--amber)}
.lg-ok .log-text{color:#34D399}
.lg-info .log-text{color:#7DD3FC}
.lg-sep .log-text{color:#38BDF8;font-weight:600}
.muted-text{color:var(--faint)!important;font-style:italic}
.pct-hl{color:var(--amber);font-weight:700}
.prompt-line .log-text{color:var(--spring)}
.term-cursor{display:inline-block;width:8px;height:14px;vertical-align:-2px;background:var(--spring);box-shadow:0 0 9px rgba(0,250,154,.8);animation:blinkRec 1.1s steps(2,start) infinite}
.jump-btn{
  position:absolute;right:18px;bottom:16px;z-index:5;
  display:inline-flex;align-items:center;gap:7px;
  font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:1.5px;color:#03130C;
  background:linear-gradient(135deg,var(--emerald),var(--spring));
  border:none;border-radius:6px;padding:9px 14px;cursor:pointer;
  box-shadow:0 6px 20px rgba(16,185,129,.4);
  opacity:0;pointer-events:none;transform:translateY(8px);transition:all .25s ease;
}
.jump-btn.visible{opacity:1;pointer-events:auto;transform:none}
.jump-btn:hover{transform:translateY(-2px)}

/* ============================================================
   04 · ARCHIVOS
   ============================================================ */
.btn-ghost{
  display:inline-flex;align-items:center;gap:8px;
  padding:9px 14px;border-radius:7px;
  font-family:var(--mono);font-size:10.5px;font-weight:700;letter-spacing:1.4px;
  color:#C4B5FD;background:rgba(139,92,246,.1);border:1px solid rgba(139,92,246,.45);
  cursor:pointer;transition:all .25s;text-transform:uppercase;
}
.btn-ghost:hover:not(:disabled){background:linear-gradient(135deg,var(--violet),#6366F1);color:#fff;border-color:transparent;box-shadow:0 6px 20px rgba(139,92,246,.4);transform:translateY(-1px)}
.btn-ghost:disabled{opacity:.6;cursor:wait}
.files-grid{display:flex;flex-direction:column;gap:12px}
.file-card{
  position:relative;display:flex;align-items:center;gap:13px;
  background:var(--panel-deep);border:1px solid var(--border);border-radius:8px;
  padding:14px 15px;transition:transform .25s,border-color .25s,box-shadow .25s;
  animation:rowIn .5s cubic-bezier(.22,1,.36,1) both;
}
.file-card:hover{transform:translateX(4px);border-color:rgba(0,250,154,.35)}
.file-card.vip{border-color:rgba(251,191,36,.4);background:linear-gradient(100deg,rgba(251,191,36,.08),rgba(139,92,246,.06)),var(--panel-deep)}
.file-ico{
  width:42px;height:42px;border-radius:8px;flex:none;font-size:19px;
  display:grid;place-items:center;color:#60A5FA;
  background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.28);
}
.file-card.vip .file-ico{color:var(--amber);background:rgba(251,191,36,.1);border-color:rgba(251,191,36,.35)}
.file-meta{display:flex;flex-direction:column;gap:2px;min-width:0;flex:1}
.file-meta strong{font-size:13.5px;font-weight:700;word-break:break-all}
.file-meta span{font-family:var(--mono);font-size:9.5px;letter-spacing:1.6px;color:var(--faint);text-transform:uppercase}
.btn-dl{
  flex:none;display:inline-flex;align-items:center;gap:7px;
  padding:9px 13px;border-radius:7px;
  font-family:var(--mono);font-size:10px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;
  color:#6EE7B7;background:rgba(16,185,129,.07);border:1px solid rgba(16,185,129,.35);
  cursor:pointer;transition:all .25s;
}
.btn-dl:hover:not(:disabled){background:linear-gradient(135deg,var(--emerald),var(--spring));color:#04150D;border-color:transparent;box-shadow:0 5px 18px rgba(16,185,129,.35)}
.btn-dl:disabled{opacity:.75;cursor:wait}
.vip-stamp{
  position:absolute;top:-9px;right:12px;
  font-family:var(--mono);font-size:8.5px;font-weight:700;letter-spacing:1.6px;
  color:var(--amber);background:#171126;
  border:1px solid rgba(251,191,36,.55);border-radius:4px;
  padding:3px 8px;transform:rotate(2deg);
}
.files-empty{
  border:1.5px dashed rgba(148,163,184,.28);border-radius:8px;
  padding:34px 20px;text-align:center;
  animation:rowIn .5s ease both;
}
.files-empty > i{font-size:24px;color:var(--dim);margin-bottom:8px;display:block}
.files-empty p{font-family:var(--serif);font-style:italic;font-size:15px;color:#B9B3A2}
.files-empty span{display:block;font-family:var(--mono);font-size:10px;letter-spacing:1.5px;color:var(--faint);margin-top:8px;text-transform:uppercase}

/* ============================================================
   05 · TEMAS DEL DÍA — Radar editorial
   ============================================================ */
.topics-updated{font-family:var(--mono);font-size:9.5px;letter-spacing:2px;color:var(--faint);text-transform:uppercase}
.topics-tabs{display:flex;gap:2px;border-bottom:1px solid var(--border);margin-bottom:4px;flex-wrap:wrap}
.ttab{
  background:none;border:none;cursor:pointer;position:relative;
  display:inline-flex;align-items:center;gap:8px;
  font-size:13px;font-weight:600;color:var(--muted);padding:11px 14px;
  transition:color .2s;
}
.ttab::after{content:'';position:absolute;left:13px;right:13px;bottom:-1px;height:2px;background:var(--spring);transform:scaleX(0);transition:transform .25s ease}
.ttab:hover{color:#E2E8F0}
.ttab.active{color:var(--paper)}
.ttab.active::after{transform:scaleX(1)}
.ttab .star{color:var(--amber);font-style:normal;font-size:12px}
.tcount{font-family:var(--mono);font-size:9.5px;font-weight:700;color:#CBD5E1;background:#33415E;border-radius:4px;padding:1.5px 6px;transition:background .25s}
.ttab.active .tcount{background:var(--spring);color:#04150D}
.topics-list{list-style:none;columns:2;column-gap:36px;min-height:130px;padding-top:6px}
.topics-list.is-empty{columns:1}
.topic-row{
  break-inside:avoid;display:flex;gap:16px;
  padding:16px 2px;border-bottom:1px dashed rgba(255,255,255,.1);
  animation:rowIn .5s cubic-bezier(.22,1,.36,1) both;animation-delay:var(--d,0ms);
}
.topic-rank{
  font-family:var(--serif);font-style:italic;font-weight:600;
  font-size:31px;line-height:1;width:42px;flex:none;
  color:transparent;-webkit-text-stroke:1px rgba(242,238,227,.5);
}
.topic-main{min-width:0}
.topic-title{font-family:var(--serif);font-size:16.5px;font-weight:600;line-height:1.35;color:var(--paper);transition:color .2s}
.topic-row:hover .topic-title{color:var(--spring)}
.topic-meta{display:flex;align-items:center;gap:11px;margin-top:9px;flex-wrap:wrap}
.tag{font-family:var(--mono);font-size:9px;letter-spacing:1.8px;padding:3px 8px;border:1px solid rgba(255,255,255,.16);border-radius:4px;color:var(--muted)}
.topic-src{font-family:var(--mono);font-size:9px;letter-spacing:1.8px;display:inline-flex;align-items:center;gap:6px;color:var(--muted)}
.topic-src::before{content:'';width:7px;height:7px;border-radius:2px}
.topic-src.et::before{background:#60A5FA}
.topic-src.pf::before{background:var(--amber)}
.heat{width:64px;height:4px;border-radius:99px;background:rgba(255,255,255,.08);overflow:hidden}
.heat i{display:block;height:100%;border-radius:99px;background:linear-gradient(90deg,var(--emerald),var(--spring))}
.heat-num{font-family:var(--mono);font-size:10px;color:var(--spring);font-weight:700}
.topic-keys{display:flex;gap:10px;margin-top:8px;flex-wrap:wrap}
.topic-keys span{font-family:var(--mono);font-size:10px;color:#5B6B84}
.topics-empty{
  padding:42px 12px;text-align:center;
  font-family:var(--serif);font-style:italic;font-size:15px;color:#8b8574;
  animation:rowIn .4s ease both;
}

/* ============================================================
   TOASTS
   ============================================================ */
.toasts{position:fixed;top:18px;right:18px;z-index:999;display:flex;flex-direction:column;gap:10px;max-width:min(370px,92vw)}
.toast{
  --tc:var(--blue);
  display:flex;align-items:flex-start;gap:11px;
  padding:13px 15px;border-radius:8px;cursor:pointer;position:relative;overflow:hidden;
  background:rgba(15,23,42,.95);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);
  border:1px solid var(--border);border-left:3px solid var(--tc);
  box-shadow:0 12px 40px rgba(0,0,0,.5);
  animation:toastIn .38s cubic-bezier(.22,1,.36,1);
}
.toast.t-success{--tc:var(--emerald)}
.toast.t-error{--tc:var(--red)}
.toast.t-warning{--tc:var(--amber)}
.toast.t-info{--tc:var(--blue)}
.toast > i{color:var(--tc);font-size:16px;margin-top:2px;flex:none}
.toast-body{display:flex;flex-direction:column;gap:2px;min-width:0}
.toast-body strong{font-size:13px}
.toast-body span{font-size:12.5px;color:var(--muted);word-break:break-word}
.toast-x{background:none;border:none;color:var(--faint);cursor:pointer;padding:2px;flex:none;transition:color .2s}
.toast-x:hover{color:#fff}
.toast::after{content:'';position:absolute;left:0;bottom:0;height:2px;width:100%;background:var(--tc);opacity:.75;animation:tbar var(--life,4200ms) linear forwards}
@keyframes tbar{to{width:0}}
.toast.out{animation:toastOut .3s ease forwards}
@keyframes toastIn{from{opacity:0;transform:translateX(40px)}}
@keyframes toastOut{to{opacity:0;transform:translateX(40px)}}
@keyframes flow{to{background-position:230% 0}}

/* ---------- Pie ---------- */
.footer{max-width:1320px;margin:0 auto;padding:30px 28px 34px;border-top:1px solid var(--border);margin-top:26px}
.f-quote{font-family:var(--serif);font-style:italic;font-size:15px;color:#B9B3A2}
.f-meta{font-family:var(--mono);font-size:9.5px;letter-spacing:2.4px;color:var(--faint);margin-top:10px;text-transform:uppercase}
.f-meta b{color:#34D399;font-weight:600}

/* ---------- Responsive ---------- */
@media (max-width:1100px){
  .grid-main,.grid-bottom{grid-template-columns:1fr}
  .topics-list{columns:1}
}
@media (max-width:720px){
  .masthead{padding:20px 16px 14px}
  .mast-right{text-align:left}
  .statusbar{padding:10px 16px}
  .deck{padding:18px 16px 6px}
  .term-body{height:280px;font-size:11.5px}
  .metrics{flex-wrap:wrap}
  .metric{min-width:33%}
  .topic-row{gap:12px}
  .topic-rank{font-size:25px;width:34px}
}
@media (prefers-reduced-motion:reduce){
  *,*::before,*::after{animation-duration:.01ms!important;animation-iteration-count:1!important;transition-duration:.01ms!important}
}
</style>
</head>
<body>

<!-- Fondo -->
<div class="bg" aria-hidden="true">
  <div class="bg-lamp"></div>
  <div class="bg-noise"></div>
  <div class="bg-rule-l"></div>
  <div class="bg-rule-r"></div>
</div>

<!-- ============ MANCHA / CABECERA EDITORIAL ============ -->
<header class="masthead">
  <div class="mast-left">
    <div class="cat-stamp" aria-hidden="true">🐈‍⬛</div>
    <div>
      <div class="mast-title">
        <h1>FLOWCATS</h1>
        <span class="ver-chip">v2.0 WEB</span>
      </div>
      <p class="mast-sub">Sala de redacción automatizada · Noticias &amp; Flowcards SEO</p>
    </div>
  </div>
  <div class="mast-right">
    <p class="mast-date" id="mastDate">—</p>
    <p class="mast-edition" id="mastEdition">EDICIÓN Nº —</p>
  </div>
</header>
<div class="double-rule" aria-hidden="true"></div>

<!-- ============ BARRA DE ESTADO ============ -->
<div class="statusbar">
  <div class="sb-left">
    <span class="ai-badge groq" id="aiBadge" role="status"><span class="pulse-dot"></span>⚡ Groq AI (Llama 3) Conectado</span>
    <span class="server-pill" id="serverPill"><span class="dot"></span>CONECTANDO…</span>
  </div>
  <div class="sb-right">
    <span class="rec" id="recBadge" hidden><i></i>REC</span>
    <span id="syncTime">SYNC --:--:--</span>
  </div>
</div>

<!-- ============ TICKER DE ÚLTIMA HORA ============ -->
<div class="ticker" aria-hidden="true">
  <div class="ticker-track" id="tickerTrack"></div>
</div>

<!-- ============ CUBIERTA ============ -->
<main class="deck">
  <div class="grid-main">

    <!-- 01 · CONFIGURACIÓN -->
    <section class="panel" aria-labelledby="cfgTag">
      <div class="panel-tag" id="cfgTag"><span class="num">01</span> Configuración de extracción</div>

      <!-- 1.1 Tipo de Proceso -->
      <div class="sec-subhead"><i class="fa-solid fa-sliders"></i> Tipo de Proceso</div>
      <div class="modes-grid">
        <label class="mode-card selected" id="modeCardFlowcards">
          <input type="checkbox" id="checkFlowcards" value="flowcards" checked>
          <span class="sc-box">
            <span style="font-size:13.5px;font-weight:700;color:var(--paper);display:flex;align-items:center;gap:6px;">
              <i class="fa-solid fa-layer-group" style="color:var(--spring)"></i> Flowcards
            </span>
            <span class="sc-sub">Máx 3 palabras clave</span>
          </span>
          <span class="sc-check"><i class="fa-solid fa-check"></i></span>
        </label>

        <label class="mode-card selected" id="modeCardTemas">
          <input type="checkbox" id="checkTemas" value="temas_del_dia" checked>
          <span class="sc-box">
            <span style="font-size:13.5px;font-weight:700;color:var(--paper);display:flex;align-items:center;gap:6px;">
              <i class="fa-solid fa-star" style="color:var(--amber)"></i> Temas del Día
            </span>
            <span class="sc-sub">Top 8 &le; 25 caracteres</span>
          </span>
          <span class="sc-check"><i class="fa-solid fa-check"></i></span>
        </label>
      </div>

      <!-- 1.2 Medio Informativo -->
      <div class="sec-subhead"><i class="fa-solid fa-newspaper"></i> Medio Informativo</div>
      <div class="sources-grid">
        <label class="source-card selected" id="sourceCardElTiempo">
          <input type="checkbox" value="El Tiempo" id="checkElTiempo" checked>
          <span class="sc-box">
            <span class="sc-mast et">EL TIEMPO</span>
            <span class="sc-sub">22 categorías · cubrimiento nacional</span>
          </span>
          <span class="sc-check"><i class="fa-solid fa-check"></i></span>
        </label>

        <label class="source-card" id="sourceCardPortafolio">
          <input type="checkbox" value="Portafolio" id="checkPortafolio">
          <span class="sc-box">
            <span class="sc-mast pf">PORTAFOLIO</span>
            <span class="sc-sub">5 grupos económicos y empresariales</span>
          </span>
          <span class="sc-check"><i class="fa-solid fa-check"></i></span>
        </label>
      </div>

      <label class="switch-row" id="ampRow">
        <span class="switch-info">
          <strong><i class="fa-solid fa-bolt-lightning"></i> Incluir URLs AMP</strong>
          <span>Google Accelerated Mobile Pages</span>
        </span>
        <span class="switch">
          <input type="checkbox" id="ampToggle" checked>
          <span class="track"><span class="thumb"></span></span>
        </span>
      </label>

      <button class="btn-run" id="runBtn">
        <i class="fa-solid fa-bolt"></i><span>Ejecutar automatización</span>
      </button>
      <p class="run-hint" id="runHint"><i class="fa-solid fa-triangle-exclamation"></i> Selecciona al menos un proceso (Flowcards/Temas) y un medio antes de ejecutar.</p>
      <p class="etiqueta">— IA: openai/gpt-oss-120b &middot; Flowcards (máx 3 palabras clave) | Temas del Día (máx 25 caracteres).</p>
    </section>

    <div class="stack">
      <!-- 02 · EJECUCIÓN -->
      <section class="panel" aria-labelledby="progTag">
        <div class="panel-tag" id="progTag">
          <span class="num">02</span> Ejecución
          <span class="chip-time"><i class="fa-solid fa-stopwatch"></i><span id="elapsed">00:00</span></span>
        </div>

        <div class="progress-top">
          <p id="actionText" role="status" aria-live="polite">Sistema en espera…</p>
          <span class="pct-num" id="pctNum">0%</span>
        </div>

        <div class="bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0" id="barWrap">
          <div class="fill" id="barFill"></div>
        </div>

        <div class="metrics">
          <div class="metric"><span class="lbl">Categorías</span><b id="mCats">0</b></div>
          <div class="metric"><span class="lbl">Archivos</span><b id="mFiles">0</b></div>
          <div class="metric"><span class="lbl">Motor IA</span><b id="mAi" class="small">Groq · Llama 3</b></div>
        </div>

        <div class="stepper" aria-hidden="true">
          <span class="step"><i class="fa-solid fa-satellite-dish"></i>EXTRACCIÓN</span><span class="st-arrow">→</span>
          <span class="step"><i class="fa-solid fa-brain"></i>ANÁLISIS IA</span><span class="st-arrow">→</span>
          <span class="step"><i class="fa-solid fa-file-excel"></i>EXCEL</span><span class="st-arrow">→</span>
          <span class="step"><i class="fa-solid fa-flag-checkered"></i>LISTO</span>
        </div>
      </section>

      <!-- 03 · CONSOLA -->
      <section class="panel p-terminal" aria-labelledby="termTag">
        <div class="panel-tag" id="termTag"><span class="num">03</span> Consola de ejecución</div>
        <div class="term-head">
          <div class="dots" aria-hidden="true"><span class="r"></span><span class="y"></span><span class="g"></span></div>
          <div class="term-title"><i class="fa-solid fa-terminal"></i>flowcats@render: ~/logs</div>
          <div class="term-actions">
            <span class="log-count" id="logCount">0 líneas</span>
            <button class="term-btn" id="copyLogsBtn"><i class="fa-solid fa-copy"></i> COPIAR LOGS</button>
            <button class="term-btn" id="clearLogsBtn"><i class="fa-solid fa-broom"></i> LIMPIAR</button>
          </div>
        </div>
        <div class="term-body" id="termBody" aria-label="Consola de logs en vivo"></div>
        <button class="jump-btn" id="jumpBtn"><i class="fa-solid fa-angles-down"></i> NUEVOS LOGS</button>
      </section>
    </div>
  </div>

  <div class="grid-bottom">
    <!-- 04 · ARCHIVOS -->
    <section class="panel" aria-labelledby="filesTag">
      <div class="panel-head-row">
        <div class="panel-tag" id="filesTag"><span class="num">04</span> Archivos generados</div>
        <button class="btn-ghost" id="downloadAllBtn" hidden>
          <i class="fa-solid fa-file-zipper"></i> Descargar todos
        </button>
      </div>
      <div id="filesZone"></div>
    </section>

    <!-- 05 · TEMAS DEL DÍA -->
    <section class="panel" aria-labelledby="topicsTag">
      <div class="panel-head-row">
        <div class="panel-tag" id="topicsTag"><span class="num">05</span> Temas del día · radar editorial</div>
        <span class="topics-updated" id="topicsUpdated">RADAR · SIN DATOS</span>
      </div>

      <div class="topics-tabs" role="tablist">
        <button class="ttab active" data-tab="top" role="tab"><i class="star">★</i> Top 8 del día <span class="tcount">0</span></button>
        <button class="ttab" data-tab="El Tiempo" role="tab">El Tiempo <span class="tcount">0</span></button>
        <button class="ttab" data-tab="Portafolio" role="tab">Portafolio <span class="tcount">0</span></button>
      </div>

      <ol class="topics-list is-empty" id="topicsList"></ol>
    </section>
  </div>
</main>

<footer class="footer">
  <p class="f-quote">“Escrito a máquina, olfateado por un gato.” — la primera regla de la casa.</p>
  <p class="f-meta"><b>Flowcats v2.0</b> · FastAPI ⚡ Render · Groq Llama 3 · © 2026 — Automatización editorial SEO</p>
</footer>

<!-- Toasts -->
<div class="toasts" id="toasts" aria-live="polite"></div>
<noscript><p style="text-align:center;padding:20px">Flowcats requiere JavaScript para funcionar.</p></noscript>

<script>
'use strict';
/* ============================================================
   FLOWCATS v2.0 · Consola web
   Contrato backend FastAPI:
     POST /api/run  ·  GET /api/status (poll 1200ms)  ·  GET /api/download/{filename}
   Compatibilidad: si el status incluye "topics", alimenta el Radar Editorial.
   ============================================================ */

/* ------------------ Configuración ------------------ */
const CONFIG = {
  POLL_MS: 1200,
  TOAST_MS: 4200,
  MAX_LOG_LINES: 600,
  FORCE_DEMO: new URLSearchParams(location.search).has('demo'),
};

/* ------------------ Utilidades ------------------ */
const $  = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];
const escapeHtml = (str) => String(str)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/'/g,'&#39;');

const els = {};
function cacheEls(){
  els.aiBadge=$('#aiBadge');      els.serverPill=$('#serverPill');
  els.recBadge=$('#recBadge');    els.syncTime=$('#syncTime');
  els.runBtn=$('#runBtn');        els.runHint=$('#runHint');
  els.amp=$('#ampToggle');        els.ampRow=$('#ampRow');
  els.pctNum=$('#pctNum');        els.barFill=$('#barFill');
  els.barWrap=$('#barWrap');      els.actionText=$('#actionText');
  els.elapsed=$('#elapsed');
  els.mCats=$('#mCats');          els.mFiles=$('#mFiles');      els.mAi=$('#mAi');
  els.termBody=$('#termBody');    els.copyBtn=$('#copyLogsBtn');
  els.clearBtn=$('#clearLogsBtn');els.jumpBtn=$('#jumpBtn');    els.logCount=$('#logCount');
  els.filesZone=$('#filesZone');  els.downloadAll=$('#downloadAllBtn');
  els.topicsList=$('#topicsList');els.topicsUpdated=$('#topicsUpdated');
  els.tickerTrack=$('#tickerTrack');
  els.mastDate=$('#mastDate');    els.mastEdition=$('#mastEdition');
  els.stepperSteps=$$('.stepper .step');
  els.sourceCards=$$('.source-card');
  els.modeCards=$$('.mode-card');
  els.checkFlowcards=$('#checkFlowcards');
  els.checkTemas=$('#checkTemas');
}

/* ------------------ Estado ------------------ */
const State = {
  running:false, progress:0,
  polling:false, pollTimer:null,
  rawLogs:[], renderedCount:0, autoFollow:true,
  files:[], startedAt:null, elapsedTimer:null,
  completionPending:false, filesLoadingShown:false,
};

/* ============================================================
   TOASTS
   ============================================================ */
const Toast = (() => {
  const ICONS = { success:'fa-circle-check', error:'fa-circle-xmark', warning:'fa-triangle-exclamation', info:'fa-circle-info' };
  function show(msg, type = 'info', title = ''){
    const wrap = document.getElementById('toasts');
    if (wrap.children.length >= 4) wrap.firstElementChild.remove();
    const t = document.createElement('div');
    t.className = `toast t-${type}`;
    t.style.setProperty('--life', CONFIG.TOAST_MS + 'ms');
    t.innerHTML = `
      <i class="fa-solid ${ICONS[type] || ICONS.info}"></i>
      <div class="toast-body">
        ${title ? `<strong>${escapeHtml(title)}</strong>` : ''}
        <span>${escapeHtml(msg)}</span>
      </div>
      <button class="toast-x" aria-label="Cerrar"><i class="fa-solid fa-xmark"></i></button>`;
    wrap.appendChild(t);
    const kill = () => { t.classList.add('out'); setTimeout(() => t.remove(), 300); };
    t.addEventListener('click', kill);
    setTimeout(kill, CONFIG.TOAST_MS);
  }
  return {
    show,
    success:(m,t)=>show(m,'success',t),
    error:(m,t)=>show(m,'error',t),
    warn:(m,t)=>show(m,'warning',t),
    info:(m,t)=>show(m,'info',t),
  };
})();

/* ============================================================
   DATOS EDITORIALES (Modo Demo)
   En producción estos temas llegan vía "topics" en /api/status.
   ============================================================ */
const TOPICS_ET = [
  {t:'Gobierno radicará la nueva reforma tributaria: qué cambia para los asalariados', c:'POLÍTICA',   h:96, k:['tributaria','congreso','impuestos']},
  {t:'Selección Colombia: nómina oficial para la doble fecha de eliminatorias',        c:'DEPORTES',   h:93, k:['selección','eliminatorias','convocatoria']},
  {t:'TRM hoy: el dólar rompe una barrera que no se veía desde 2021',                   c:'ECONOMÍA',   h:90, k:['dólar','trm','peso']},
  {t:'Ideam activa alerta naranja por lluvias en la región Andina',                    c:'NACIÓN',     h:87, k:['ideam','lluvias','alerta']},
  {t:'Casos de dengue se duplican: así avanza el plan de vacunación',                  c:'SALUD',      h:82, k:['dengue','vacunación','minsalud']},
  {t:'Boom de IA generativa en universidades: el debate que divide a los rectores',    c:'TECNOLOGÍA', h:79, k:['ia','universidades','educación']},
  {t:'Festival Estéreo Picnic confirma sus cabezas de cartel para 2027',               c:'CULTURA',    h:74, k:['estéreo picnic','conciertos','bogotá']},
  {t:'Movilidad en Bogotá: así operarán los cierres por las obras de la Caracas',      c:'BOGOTÁ',     h:71, k:['movilidad','caracas','transmilenio']},
].map(x => ({ ...x, s:'El Tiempo' }));

const TOPICS_PF = [
  {t:'Resultados de Bancolombia superan las previsiones de los analistas',             c:'EMPRESAS',  h:94, k:['bancolombia','utilidades','banca']},
  {t:'Ecopetrol ajusta su plan de transición energética al 2030',                      c:'ECONOMÍA',  h:90, k:['ecopetrol','energía','transición']},
  {t:'Startups colombianas levantan US$120 millones en ronda regional',                c:'NEGOCIOS',  h:85, k:['startups','venture capital','rondas']},
  {t:'Subsidios VIS: los cambios que deben conocer los compradores',                   c:'FINANZAS',  h:81, k:['vivienda','vis','subsidios']},
  {t:'Exportaciones de café crecen 18% impulsadas por Asia',                           c:'MERCADOS',  h:77, k:['café','exportaciones','asia']},
  {t:'El Dorado inaugura su nueva terminal de carga aérea',                            c:'EMPRESAS',  h:72, k:['el dorado','logística','carga']},
].map(x => ({ ...x, s:'Portafolio' }));

const TOPICS_TOP = [
  TOPICS_ET[0], TOPICS_PF[0], TOPICS_ET[1], TOPICS_PF[1],
  TOPICS_ET[2], TOPICS_ET[3], TOPICS_PF[2], TOPICS_ET[4],
];

/* ============================================================
   MOCK SERVER · simulación del backend FastAPI (?demo o fallback)
   ============================================================ */
const MockServer = (() => {
  const ET_CATS = ['Nación','Internacional','Política','Justicia','Economía','Deportes','Tecnología',
    'Salud','Ciencia','Cultura','Educación','Entretenimiento','Medio Ambiente','Gastronomía','Turismo',
    'Moda','Motor','Vivienda','Trabajo','Opinión','Investigación','Bogotá'];
  const PF_GROUPS = ['Economía','Empresas','Negocios','Finanzas','Mercados'];

  let snap = {
    running:false, progress:0,
    status:'Sistema en espera · listo para ejecutar',
    groq_active:true, files:[], topics:{},
    logs:['[✓] Flowcats v2.0 inicializado en Render.','[*] Esperando instrucciones del operador...'],
  };
  let timers = [];
  const rand = (a,b) => Math.floor(a + Math.random() * (b - a));

  function run(payload){
    if (snap.running){
      return Promise.reject(Object.assign(new Error('Ya hay un proceso en ejecución.'), { status:400 }));
    }
    timers.forEach(clearTimeout); timers = [];
    snap = {
      running:true, progress:0, status:'Inicializando workers en la nube...',
      groq_active: Math.random() > 0.12, files:[],
      topics:{ top:[], 'El Tiempo':[], 'Portafolio':[] },
      logs:[],
    };
    const events = buildEvents(payload);
    let t = 200;
    for (const ev of events){
      t += ev.wait;
      timers.push(setTimeout(() => applyEvent(ev, payload), t));
    }
    return Promise.resolve({ message:'Proceso iniciado en segundo plano', sources:payload.selected_sources });
  }

  function applyEvent(ev, payload){
    snap.logs.push(ev.log);
    if (ev.p != null) snap.progress = ev.p;
    if (ev.action) snap.status = ev.action;
    if (ev.tpKey) snap.topics[ev.tpKey] = ev.tpVal;
    if (ev.done){
      snap.running = false; snap.progress = 100;
      snap.status = 'Proceso completado exitosamente';
      snap.files = [...payload.selected_sources.map(s => `${s}.xlsx`), 'TEMAS DEL DÍA.xlsx'];
    }
  }

  function buildEvents(payload){
    const sources = payload.selected_sources || [];
    const amp = !!payload.include_amp;
    const jobs = [];
    if (sources.includes('El Tiempo'))  jobs.push({ src:'El Tiempo',  items:ET_CATS });
    if (sources.includes('Portafolio')) jobs.push({ src:'Portafolio', items:PF_GROUPS });

    const totalItems = jobs.reduce((a,j) => a + j.items.length, 0) || 1;
    const units = totalItems + (amp ? jobs.length : 0);
    let done = 0;
    const nextP = () => Math.min(88, 8 + Math.round(80 * (++done / units)));
    const ev = [];

    ev.push({ log:'=== INICIANDO PROCESO EN RENDER WEB ===', p:2, action:'Inicializando workers en la nube...', wait:rand(350,550) });
    ev.push({ log:'[*] Autenticando con Groq Cloud (Llama 3 70B)...', p:4, wait:rand(380,620) });
    ev.push({ log: snap.groq_active ? '[✓] LLM verificado · groq_active=true' : '[!] Groq no disponible · fallback a motor heurístico',
              p:6, action:'Motor IA verificado', wait:rand(300,520) });

    jobs.forEach(job => {
      ev.push({ log:`── Extrayendo fuente: ${job.src} ──`, action:`Extrayendo ${job.src}...`, wait:rand(260,420) });
      job.items.forEach((cat, i) => {
        const p = nextP();
        ev.push({ log:`[*] [${i + 1}/${job.items.length}] ${cat} procesado (${p}%)...`,
                  p, action:`Procesando "${cat}" · ${job.src} (${p}%)`, wait:rand(240,560) });
        if (job.items.length > 10 && i === 8){
          ev.push({ log:`[!] Timeout 408 en "${cat}" · reintento automático (1/3)...`, wait:rand(320,520) });
          ev.push({ log:'[✓] Respuesta recuperada tras reintento', wait:rand(240,420) });
        }
      });
      if (amp){
        const ok = rand(11, 18);
        ev.push({ log:`[*] Validando URLs AMP de ${job.src} (Google Accelerated Mobile Pages)...`,
                  action:`Validando URLs AMP · ${job.src}`, p:nextP(), wait:rand(300,520) });
        ev.push({ log:`[✓] ${ok}/${ok} URLs AMP verificadas correctamente`, wait:rand(240,420) });
      }
      ev.push({ log:`[+] Archivo ${job.src}.xlsx generado con ${job.items.length * rand(3,5)} flowcards SEO.`,
                action:`Empaquetando ${job.src}.xlsx`,
                tpKey:job.src, tpVal:(job.src === 'El Tiempo' ? TOPICS_ET : TOPICS_PF),
                wait:rand(360,560) });
    });

    ev.push({ log:'[*] Sintetizando tendencias del día con IA...', p:92, action:'Generando TEMAS DEL DÍA...', wait:rand(520,820) });
    ev.push({ log:'[+] Archivo TEMAS DEL DÍA.xlsx generado con 8 temas clave.', p:98,
              action:'Levantando el radar editorial...', tpKey:'top', tpVal:TOPICS_TOP, wait:rand(420,640) });
    ev.push({ log:'=== PROCESO COMPLETADO EXITOSAMENTE ===', p:100, action:'Proceso completado exitosamente', done:true, wait:320 });
    return ev;
  }

  function status(){
    return Promise.resolve({
      running:snap.running, progress:snap.progress, status:snap.status,
      groq_active:snap.groq_active, files:[...snap.files], logs:[...snap.logs],
      topics:{ top:[...(snap.topics.top||[])], 'El Tiempo':[...(snap.topics['El Tiempo']||[])], 'Portafolio':[...(snap.topics['Portafolio']||[])] },
    });
  }

  function download(name){
    const content = `FLOWCATS · ARCHIVO DEMO\n========================\nArchivo: ${name}\nGenerado: ${new Date().toLocaleString('es-CO')}\n(Contenido de demostración — conecta el backend FastAPI para archivos reales)`;
    return new Promise(res => setTimeout(() =>
      res(new Blob([content], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })), 350));
  }

  return { run, status, download };
})();

/* ============================================================
   CAPA DE API
   ============================================================ */
const API = (() => {
  let mock = CONFIG.FORCE_DEMO;
  const isMock = () => mock;
  const enableMock = () => { mock = true; };

  async function status(){
    if (mock) return MockServer.status();
    const r = await fetch('/api/status', { cache:'no-store' });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    return r.json();
  }

  async function run(payload){
    if (mock) return MockServer.run(payload);
    const r = await fetch('/api/run', {
      method:'POST',
      headers:{ 'Content-Type':'application/json' },
      body: JSON.stringify(payload),
    });
    let data = {};
    try { data = await r.json(); } catch (_) {}
    if (!r.ok){
      const e = new Error(data.detail || `Error del servidor (${r.status})`);
      e.status = r.status;
      throw e;
    }
    return data;
  }

  async function download(name){
    if (mock) return MockServer.download(name);
    const r = await fetch(`/api/download/${encodeURIComponent(name)}`, { cache:'no-store' });
    if (!r.ok) throw new Error(`No se pudo descargar "${name}" (HTTP ${r.status})`);
    return r.blob();
  }

  return { status, run, download, isMock, enableMock };
})();

/* ============================================================
   TICKER de última hora
   ============================================================ */
const Ticker = (() => {
  function set(items){
    const arr = (items && items.length) ? items : [
      { c:'SISTEMA', t:'En espera de la primera extracción — el radar editorial se llenará solo' },
      { c:'FLOWCATS', t:'FastAPI · Render · Groq Llama 3 · v2.0 Web' },
    ];
    const html = arr.map(it =>
      `<span class="tk"><b>${escapeHtml(it.c)}</b>${escapeHtml(it.t)}<span class="tk-dot">◆</span></span>`).join('');
    els.tickerTrack.innerHTML =
      `<div class="tk-group">${html}</div><div class="tk-group" aria-hidden="true">${html}</div>`;
  }
  return { set };
})();

/* ============================================================
   05 · RADAR EDITORIAL (Temas del día)
   ============================================================ */
const Topics = (() => {
  const TABS = ['top', 'El Tiempo', 'Portafolio'];
  let store = { top:[], 'El Tiempo':[], 'Portafolio':[] };
  let active = 'top';
  const pad2 = n => String(n).padStart(2, '0');

  function renderTabs(){
    $$('.ttab').forEach(b => {
      const key = b.dataset.tab;
      b.classList.toggle('active', key === active);
      b.querySelector('.tcount').textContent = (store[key] || []).length;
    });
  }

  function renderList(){
    const list = store[active] || [];
    const ul = els.topicsList;
    ul.classList.toggle('is-empty', !list.length);
    if (!list.length){
      ul.innerHTML = `<li class="topics-empty">${State.running
        ? 'Radar sintonizando… los temas aterrizan aquí mientras el gato redacta.'
        : 'Ejecuta la automatización para levantar el radar editorial de hoy.'}</li>`;
      return;
    }
    ul.innerHTML = list.map((t, i) => {
      const cls = (t.s === 'Portafolio') ? 'pf' : 'et';
      const keys = (t.k || []).map(k => `<span>#${escapeHtml(k)}</span>`).join('');
      return `<li class="topic-row" style="--d:${i * 70}ms">
        <span class="topic-rank">${pad2(i + 1)}</span>
        <div class="topic-main">
          <h4 class="topic-title">${escapeHtml(t.t)}</h4>
          <div class="topic-meta">
            <span class="tag">${escapeHtml(t.c || '—')}</span>
            <span class="topic-src ${cls}">${escapeHtml(t.s || '')}</span>
            <span class="heat"><i style="width:${Math.min(100, t.h || 0)}%"></i></span>
            <span class="heat-num">${t.h || 0}</span>
          </div>
          ${keys ? `<div class="topic-keys">${keys}</div>` : ''}
        </div>
      </li>`;
    }).join('');
  }

  function tickerItems(){
    const src = store.top.length ? store.top
              : store['El Tiempo'].length ? store['El Tiempo']
              : store['Portafolio'];
    return src.map(t => ({ c:t.c, t:t.t }));
  }

  /** Recibe s.topics desde /api/status (si el backend lo incluye) */
  function update(obj){
    if (!obj || typeof obj !== 'object') return;
    let changed = false;
    TABS.forEach(k => {
      if (Array.isArray(obj[k]) && obj[k].length !== (store[k] || []).length){
        store[k] = obj[k];
        changed = true;
      }
    });
    if (changed){
      renderTabs();
      renderList();
      els.topicsUpdated.textContent = 'RADAR · ' + new Date().toLocaleTimeString('es-CO', { hour12:false });
      Ticker.set(tickerItems());
    }
  }

  function reset(){
    store = { top:[], 'El Tiempo':[], 'Portafolio':[] };
    renderTabs();
    renderList();
    els.topicsUpdated.textContent = 'RADAR · SIN DATOS';
    Ticker.set([]);
  }

  function bind(){
    $$('.ttab').forEach(b => b.addEventListener('click', () => {
      active = b.dataset.tab;
      renderTabs();
      renderList();
    }));
  }

  return { update, reset, bind };
})();

/* ============================================================
   UI · estado, progreso, métricas
   ============================================================ */
function updateServerPill(){
  els.serverPill.innerHTML = API.isMock()
    ? '<span class="dot amber"></span>MODO DEMO LOCAL'
    : '<span class="dot green"></span>RENDER · API CONECTADA';
}

function updateAiBadge(active){
  if (active){
    els.aiBadge.className = 'ai-badge groq';
    els.aiBadge.innerHTML = '<span class="pulse-dot"></span>⚡ Groq AI (Llama 3) Conectado';
    els.mAi.textContent = 'Groq · Llama 3';
  } else {
    els.aiBadge.className = 'ai-badge heur';
    els.aiBadge.innerHTML = '<span class="pulse-dot amber"></span>⚙️ Modo Heurístico';
    els.mAi.textContent = 'Heurístico';
  }
}

let shownPct = 0, pctRaf = null;
function setProgress(target){
  target = Math.max(0, Math.min(100, Math.round(target)));
  State.progress = target;
  els.barFill.style.width = target + '%';
  els.barWrap.setAttribute('aria-valuenow', target);
  const from = shownPct, start = performance.now(), dur = 480;
  cancelAnimationFrame(pctRaf);
  const stepFn = (now) => {
    const k = Math.min(1, (now - start) / dur);
    const eased = 1 - Math.pow(1 - k, 3);
    els.pctNum.textContent = Math.round(from + (target - from) * eased) + '%';
    if (k < 1) pctRaf = requestAnimationFrame(stepFn);
  };
  pctRaf = requestAnimationFrame(stepFn);
  shownPct = target;
  updateStepper(target);
}

function updateStepper(p){
  const th = [1, 40, 78, 100];
  els.stepperSteps.forEach((el, i) => {
    const reached = p >= th[i];
    const completed = i < th.length - 1 ? p >= th[i + 1] : p >= 100;
    el.classList.toggle('active', reached && !completed);
    el.classList.toggle('done', reached && completed);
  });
}

function fmtElapsed(ms){
  const s = Math.floor(ms / 1000), m = Math.floor(s / 60), h = Math.floor(m / 60);
  const pad = n => String(n).padStart(2, '0');
  return h ? `${pad(h)}:${pad(m % 60)}:${pad(s % 60)}` : `${pad(m)}:${pad(s % 60)}`;
}
function startElapsed(){
  stopElapsed();
  State.startedAt = Date.now();
  els.elapsed.textContent = '00:00';
  State.elapsedTimer = setInterval(() => {
    els.elapsed.textContent = fmtElapsed(Date.now() - State.startedAt);
  }, 1000);
}
function stopElapsed(){
  if (State.elapsedTimer){ clearInterval(State.elapsedTimer); State.elapsedTimer = null; }
}

/* ============================================================
   UI · Terminal
   ============================================================ */
let promptLine = null;

function buildTerminal(){
  promptLine = document.createElement('div');
  promptLine.className = 'term-line prompt-line';
  promptLine.innerHTML = '<span class="log-time">$</span><span class="log-text"><span class="term-cursor"></span></span>';
  els.termBody.appendChild(promptLine);
  ensurePlaceholder();
}

function ensurePlaceholder(){
  if (els.termBody.querySelector('.placeholder-line')) return;
  const d = document.createElement('div');
  d.className = 'term-line placeholder-line';
  d.innerHTML = '<span class="log-time">--:--:--</span><span class="log-text muted-text">Sistema listo. Ejecuta la automatización para ver los logs en vivo…</span>';
  els.termBody.insertBefore(d, promptLine);
}

function clearTerm(){
  $$('.term-line', els.termBody).forEach(n => { if (n !== promptLine) n.remove(); });
  els.jumpBtn.classList.remove('visible');
}

function classify(line){
  if (/^\s*===/.test(line)) return 'lg-head';
  if (line.startsWith('[ERROR]')) return 'lg-error';
  if (line.startsWith('[!]')) return 'lg-warn';
  if (line.startsWith('[+]') || line.startsWith('[✓]')) return 'lg-ok';
  if (line.startsWith('[*]')) return 'lg-info';
  if (line.startsWith('──')) return 'lg-sep';
  return 'lg-plain';
}

function buildLogLine(line){
  const row = document.createElement('div');
  row.className = 'term-line ' + classify(line);
  const time = document.createElement('span');
  time.className = 'log-time';
  time.textContent = new Date().toLocaleTimeString('es-CO', { hour12:false });
  const txt = document.createElement('span');
  txt.className = 'log-text';
  txt.innerHTML = escapeHtml(line).replace(/(\d+(?:[.,]\d+)?\s*%)/g, '<span class="pct-hl">$1</span>');
  row.append(time, txt);
  return row;
}

function appendLogs(logs){
  if (!Array.isArray(logs)) return;
  if (logs.length < State.renderedCount){ State.renderedCount = 0; clearTerm(); }
  State.rawLogs = logs;
  const fresh = logs.slice(State.renderedCount);
  if (!fresh.length) return;

  const ph = els.termBody.querySelector('.placeholder-line');
  if (ph) ph.remove();

  const frag = document.createDocumentFragment();
  fresh.forEach(l => frag.appendChild(buildLogLine(l)));
  els.termBody.insertBefore(frag, promptLine);
  State.renderedCount = logs.length;

  const rows = $$('.term-line', els.termBody);
  if (rows.length > CONFIG.MAX_LOG_LINES){
    rows.slice(0, rows.length - CONFIG.MAX_LOG_LINES).forEach(r => { if (r !== promptLine) r.remove(); });
  }

  els.logCount.textContent = `${logs.length} líneas`;
  if (State.autoFollow){
    els.termBody.scrollTop = els.termBody.scrollHeight;
  } else {
    els.jumpBtn.classList.add('visible');
  }
}

async function copyLogs(){
  const text = State.rawLogs.join('\n');
  if (!text){ Toast.warn('No hay logs para copiar.'); return; }
  try {
    await navigator.clipboard.writeText(text);
  } catch (_) {
    const ta = document.createElement('textarea');
    ta.value = text; document.body.appendChild(ta);
    ta.select(); document.execCommand('copy'); ta.remove();
  }
  Toast.success(`${State.rawLogs.length} líneas copiadas al portapapeles.`);
}

function clearLogs(){
  clearTerm();
  State.renderedCount = State.rawLogs.length;
  els.logCount.textContent = '0 líneas';
  ensurePlaceholder();
  Toast.info('Consola limpiada.');
}

/* ============================================================
   UI · Archivos / descargas
   ============================================================ */
function renderFilesEmpty(){
  els.filesZone.innerHTML = `
    <div class="files-empty">
      <i class="fa-regular fa-folder-open"></i>
      <p>Todavía no huele a papel recién impreso.</p>
      <span>Los XLSX llegan cuando la barra marca 100%</span>
    </div>`;
  els.downloadAll.hidden = true;
}

function renderFilesLoading(){
  els.filesZone.innerHTML = `
    <div class="files-empty">
      <span class="spinner big"></span>
      <p>Imprimiendo los libros Excel en la nube…</p>
      <span>Disponibles al completar el proceso</span>
    </div>`;
  els.downloadAll.hidden = true;
}

function buildFileCard(name, i){
  const vip = /TEMAS/i.test(name);
  const card = document.createElement('article');
  card.className = 'file-card' + (vip ? ' vip' : '');
  card.style.animationDelay = `${i * 110}ms`;
  card.innerHTML = `
    ${vip ? '<span class="vip-stamp">★ TOP 8 · VIP</span>' : ''}
    <div class="file-ico"><i class="fa-solid ${vip ? 'fa-star' : 'fa-file-excel'}"></i></div>
    <div class="file-meta">
      <strong>${escapeHtml(name)}</strong>
      <span>${vip ? 'XLSX · 8 temas clave del día' : 'XLSX · Flowcards SEO'}</span>
    </div>
    <button class="btn-dl"><i class="fa-solid fa-download"></i> Descargar</button>`;
  card.querySelector('.btn-dl').addEventListener('click', (e) => downloadFile(name, e.currentTarget));
  return card;
}

function renderFiles(files){
  State.files = files || [];
  els.mFiles.textContent = State.files.length;
  if (!State.files.length){ renderFilesEmpty(); return; }
  els.filesZone.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'files-grid';
  State.files.forEach((name, i) => grid.appendChild(buildFileCard(name, i)));
  els.filesZone.appendChild(grid);
  els.downloadAll.hidden = State.files.length < 2;
}

function triggerBlob(blob, name){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function setDlLoading(btn, on){
  if (on){
    btn.dataset.orig = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner sm"></span> Bajando…';
  } else if (btn.dataset.orig){
    btn.innerHTML = btn.dataset.orig;
    btn.disabled = false;
  }
}

async function downloadFile(name, btn){
  if (btn) setDlLoading(btn, true);
  try {
    const blob = await API.download(name);
    triggerBlob(blob, name);
    Toast.success(`Descargando ${name}…`, '📥 Descarga iniciada');
  } catch (err){
    Toast.error(err.message || 'Error al descargar el archivo.');
  } finally {
    if (btn) setDlLoading(btn, false);
  }
}

async function downloadAll(){
  const btn = els.downloadAll;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner sm"></span> Preparando…';
  for (const f of State.files){
    await downloadFile(f);
    await new Promise(r => setTimeout(r, 350));
  }
  btn.disabled = false;
  btn.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Descargar todos';
  Toast.success('Se solicitaron todos los archivos generados.', '✅ Lote completo');
}

/* ============================================================
   UI · Botón principal, validación y bloqueo
   ============================================================ */
const getSelectedSources = () =>
  els.sourceCards.map(c => c.querySelector('input')).filter(i => i.checked).map(i => i.value);

const getSelectedModes = () =>
  els.modeCards.map(c => c.querySelector('input')).filter(i => i.checked).map(i => i.value);

function getProcessType(){
  const modes = getSelectedModes();
  if (modes.includes('flowcards') && modes.includes('temas_del_dia')) return 'both';
  if (modes.includes('flowcards')) return 'flowcards';
  if (modes.includes('temas_del_dia')) return 'temas_del_dia';
  return 'both';
}

function syncRunControls(){
  const sources = getSelectedSources();
  const modes = getSelectedModes();
  const isValid = sources.length > 0 && modes.length > 0;
  els.runBtn.disabled = State.running || !isValid;
  els.runHint.classList.toggle('show', !State.running && !isValid);
}

function setRunningLock(on){
  els.sourceCards.forEach(c => {
    c.querySelector('input').disabled = on;
    c.classList.toggle('locked', on);
  });
  els.modeCards.forEach(c => {
    c.querySelector('input').disabled = on;
    c.classList.toggle('locked', on);
  });
  els.amp.disabled = on;
  els.ampRow.classList.toggle('locked', on);
}

function setRunButton(running){
  if (running){
    els.runBtn.classList.add('loading');
    els.runBtn.innerHTML = '<span class="spinner"></span><span>Extrayendo datos…</span>';
  } else {
    els.runBtn.classList.remove('loading');
    els.runBtn.innerHTML = '<i class="fa-solid fa-bolt"></i><span>Ejecutar automatización</span>';
  }
  setRunningLock(running);
  syncRunControls();
}

function attachRipple(el){
  el.addEventListener('pointerdown', e => {
    if (el.disabled) return;
    const rect = el.getBoundingClientRect();
    const d = Math.max(rect.width, rect.height);
    const span = document.createElement('span');
    span.className = 'ripple';
    span.style.width = span.style.height = d + 'px';
    span.style.left = (e.clientX - rect.left - d / 2) + 'px';
    span.style.top  = (e.clientY - rect.top  - d / 2) + 'px';
    el.appendChild(span);
    span.addEventListener('animationend', () => span.remove());
  });
}

/* ============================================================
   POLLING · cada 1200ms solo con tarea activa
   ============================================================ */
function startPolling(){
  if (State.polling) return;
  State.polling = true;
  pollTick();
}
function stopPolling(){
  State.polling = false;
  if (State.pollTimer){ clearTimeout(State.pollTimer); State.pollTimer = null; }
}
async function pollTick(){
  if (!State.polling) return;
  try {
    applyStatus(await API.status());
  } catch (_) {
    Toast.error('No se pudo consultar /api/status. Reintentando…');
  }
  if (State.polling) State.pollTimer = setTimeout(pollTick, CONFIG.POLL_MS);
}

/* ============================================================
   SINCRONIZACIÓN DE ESTADO DESDE EL SERVIDOR
   ============================================================ */
function applyStatus(s){
  const progress = typeof s.progress === 'number' ? s.progress : State.progress;
  State.running = !!s.running;

  setProgress(progress);
  if (s.status) els.actionText.textContent = s.status;
  updateAiBadge(s.groq_active !== false);
  if (Array.isArray(s.logs)) appendLogs(s.logs);
  if (s.topics) Topics.update(s.topics);
  els.mCats.textContent = State.rawLogs.filter(l => /^\[\*\] \[\d+\//.test(l)).length;

  /* Indicadores de barra de estado */
  els.recBadge.hidden = !State.running;
  els.syncTime.textContent = 'SYNC ' + new Date().toLocaleTimeString('es-CO', { hour12:false });

  if (State.running){
    if (!els.runBtn.classList.contains('loading')) setRunButton(true);
    if (!State.filesLoadingShown){ renderFilesLoading(); State.filesLoadingShown = true; }
    if (!State.elapsedTimer) startElapsed();
    if (!State.polling) startPolling();
  } else {
    stopPolling();
    stopElapsed();
    setRunButton(false);

    if (progress >= 100 && Array.isArray(s.files) && s.files.length){
      if (State.files.join('|') !== s.files.join('|')) renderFiles(s.files);
      setProgress(100);
      els.actionText.textContent = s.status || 'Proceso completado exitosamente';
      if (State.completionPending){
        Toast.success(s.status || 'Proceso completado exitosamente.', '✅ Edición lista');
        State.completionPending = false;
      }
    } else if (!State.files.length){
      renderFilesEmpty();
    }
  }
}

/* ============================================================
   EJECUCIÓN DEL PIPELINE
   ============================================================ */
async function runFlow(){
  if (State.running) return;
  const sources = getSelectedSources();
  if (!sources.length){
    Toast.warn('Selecciona al menos un proceso (Flowcards/Temas) y un medio antes de ejecutar.');
    return;
  }

  setRunButton(true);
  els.actionText.textContent = 'Enviando la orden a la sala de máquinas…';

  try {
    const res = await API.run({
      selected_sources: sources,
      process_type: getProcessType(),
      include_amp: els.amp.checked,
    });

    /* Reset de la corrida */
    State.rawLogs = []; State.renderedCount = 0;
    clearTerm(); ensurePlaceholder();
    els.logCount.textContent = '0 líneas';
    State.files = []; State.filesLoadingShown = false;
    State.completionPending = true;
    State.running = true;
    renderFilesLoading(); State.filesLoadingShown = true;
    Topics.reset();
    startElapsed();

    Toast.success(res.message || 'Proceso iniciado en segundo plano.', '🚀 Automatización iniciada');
    startPolling();
  } catch (err){
    State.completionPending = false;
    setRunButton(false);
    Toast.error(err.message || 'No se pudo iniciar el proceso.', '⚠ Error');
  }
}

/* ============================================================
   EVENTOS E INICIALIZACIÓN
   ============================================================ */
function bindEvents(){
  els.runBtn.addEventListener('click', runFlow);
  els.copyBtn.addEventListener('click', copyLogs);
  els.clearBtn.addEventListener('click', clearLogs);
  els.downloadAll.addEventListener('click', downloadAll);

  els.sourceCards.forEach(card => {
    const input = card.querySelector('input');
    input.addEventListener('change', () => {
      card.classList.toggle('selected', input.checked);
      syncRunControls();
    });
  });

  els.modeCards.forEach(card => {
    const input = card.querySelector('input');
    input.addEventListener('change', () => {
      card.classList.toggle('selected', input.checked);
      syncRunControls();
    });
  });

  /* Auto-scroll inteligente de la terminal */
  els.termBody.addEventListener('scroll', () => {
    const nearBottom = els.termBody.scrollTop + els.termBody.clientHeight >= els.termBody.scrollHeight - 40;
    State.autoFollow = nearBottom;
    if (nearBottom) els.jumpBtn.classList.remove('visible');
  });
  els.jumpBtn.addEventListener('click', () => {
    State.autoFollow = true;
    els.termBody.scrollTo({ top: els.termBody.scrollHeight, behavior:'smooth' });
    els.jumpBtn.classList.remove('visible');
  });

  /* Atajo Ctrl/Cmd + Enter */
  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter' && !els.runBtn.disabled) runFlow();
  });
}

/* Cabecera editorial: fecha del día + número de edición */
function setMasthead(){
  const now = new Date();
  const txt = now.toLocaleDateString('es-CO', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
  els.mastDate.textContent = txt.charAt(0).toUpperCase() + txt.slice(1);
  const edition = Math.max(1, Math.floor((now - new Date('2025-11-01T00:00:00')) / 864e5));
  els.mastEdition.textContent = `EDICIÓN Nº ${edition} · RENDER / FASTAPI`;
}

/* Aparición progresiva de paneles al hacer scroll */
function initReveal(){
  const io = new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting){
        e.target.classList.add('in');
        io.unobserve(e.target);
      }
    });
  }, { threshold:.1 });
  $$('.panel').forEach(p => io.observe(p));
}

document.addEventListener('DOMContentLoaded', async () => {
  cacheEls();
  buildTerminal();
  bindEvents();
  setMasthead();
  attachRipple(els.runBtn);
  attachRipple(els.downloadAll);
  Topics.bind();
  Topics.reset();
  Ticker.set([]);
  setProgress(0);
  renderFilesEmpty();
  syncRunControls();
  initReveal();

  if (API.isMock()){
    updateServerPill();
    Toast.info('Backend simulado localmente. Añade ?demo a la URL para forzar este modo.', '🧪 Modo Demo');
  }

  /* Sincronización inicial: retomamos el hilo si el servidor ya estaba procesando */
  try {
    applyStatus(await API.status());
    if (!API.isMock()) updateServerPill();
    if (State.running){ State.completionPending = true; startElapsed(); }
  } catch (_) {
    API.enableMock();
    updateServerPill();
    Toast.warn('Backend FastAPI no detectado. Activando Modo Demo local.', '🔌 Sin conexión');
    applyStatus(await MockServer.status());
  }
});
</script>
</body>
</html>
'''


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"[OK] Iniciando Flowcats v2.0 Web en puerto {port}...")
    uvicorn.run("app:app", host="0.0.0.0", port=port, reload=False)
