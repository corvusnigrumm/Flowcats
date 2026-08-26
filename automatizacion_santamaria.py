# -*- coding: utf-8 -*-
"""
automatizacion_santamaria.py
====================
Extrae UN articulo reciente por categoria de https://www.eltiempo.com/ y 
https://www.portafolio.co/ y los exporta a libros Excel separados:
    - El Tiempo.xlsx
    - Portafolio.xlsx

Caracteristicas:
    - Titulo corto inteligente (max 2 palabras con sentido coherente)
    - Casilla de verificacion que resalta la fila en verde (solo admite ✔)
    - Diseño en fuente Century Gothic con encabezados de fondo negro.
    - Filtro de noticias exclusivas para suscriptores (solo El Tiempo)
    - Categorias agrupadas para Portafolio (5 grupos)
    - Columna de resumen corto al final

Dependencias:
    pip install requests openpyxl lxml

Uso:
    python automatizacion_santamaria.py
"""

import sys
import io
import re
import os
import json
from datetime import datetime
from email.utils import parsedate_to_datetime

from typing import Dict, TypedDict

import requests  # type: ignore
import openpyxl  # type: ignore
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # type: ignore
from openpyxl.formatting.rule import FormulaRule  # type: ignore
from openpyxl.utils import get_column_letter  # type: ignore
from openpyxl.worksheet.datavalidation import DataValidation  # type: ignore

try:
    from lxml import etree as ET  # type: ignore
except ImportError:
    import xml.etree.ElementTree as ET

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

# ─────────────────────────────────────────────
# FEEDS RSS Y CONFIGURACION DE SITIOS
# ─────────────────────────────────────────────

class SiteConfig(TypedDict):
    output_file: str
    url_base: str
    feeds: Dict[str, str]
    # feeds_grouped es alternativo: Dict[str, list[str]] para categorias agrupadas

# Definición de los dos diarios a extraer
SITES: Dict[str, SiteConfig] = {
    "El Tiempo": {
        "output_file": "El Tiempo.xlsx",
        "url_base": "eltiempo.com",
        "feeds": {
            "Bogota":   "https://www.eltiempo.com/rss/bogota.xml",
            "Colombia": "https://www.eltiempo.com/rss/colombia.xml",
            "Mundo":    "https://www.eltiempo.com/rss/mundo.xml",
            "Economia": "https://www.eltiempo.com/rss/economia.xml",
            "Deportes": "https://www.eltiempo.com/rss/deportes.xml",
            "Politica": "https://www.eltiempo.com/rss/politica.xml",
            "Cultura":  "https://www.eltiempo.com/rss/cultura.xml",
            "Vida":     "https://www.eltiempo.com/rss/vida.xml",
            "Justicia": "https://www.eltiempo.com/rss/justicia.xml",
            "Salud":    "https://www.eltiempo.com/rss/salud.xml",
        }
    },
}

TREND_KEYWORDS = {
    "crisis", "reforma", "elecciones", "inflacion", "inflacion", "dolar", "dolar",
    "subsidio", "impuestos", "pension", "pension", "salario", "empleo", "desempleo",
    "salud", "seguridad", "justicia", "guerra", "arancel", "precio", "precios",
    "petroleo", "petroleo", "tecnologia", "tecnologia", "inteligencia artificial",
    "ia", "seleccion", "seleccion", "futbol", "futbol", "escandalo", "escandalo",
    "historico", "historico", "record", "record", "viral", "denuncia"
}

SEO_KEYWORDS = {
    "como", "como", "que", "que", "por que", "por que", "claves", "guia", "guia",
    "paso a paso", "explicamos", "esto significa", "fechas", "requisitos",
    "quienes", "quienes", "cuando", "cuando", "precio", "precios", "ranking"
}

CATEGORY_POTENTIAL = {
    "Colombia": 12,
    "Politica": 12,
    "Economia": 11,
    "Salud": 11,
    "Bogota": 9,
    "Mundo": 8,
    "Justicia": 8,
    "Vida": 7,
    "Deportes": 7,
    "Cultura": 5,
}

# Portafolio usa feeds agrupados: cada categoria muestra una sola etiqueta 
# pero recoge articulos de multiples feeds RSS.
PORTAFOLIO_CONFIG = {
    "output_file": "Portafolio.xlsx",
    "url_base": "portafolio.co",
    "feeds_grouped": {
        "Economía, indicadores": [
            "https://www.portafolio.co/rss/economia.xml",
            "https://www.portafolio.co/rss/indicadores-economicos.xml",
        ],
        "Tendencias": [
            "https://www.portafolio.co/rss/tendencias.xml",
        ],
        "Internacional, emprendimiento, tecno": [
            "https://www.portafolio.co/rss/internacional.xml",
            "https://www.portafolio.co/rss/negocios/emprendimiento.xml",
            "https://www.portafolio.co/rss/innovacion.xml",
        ],
        "Energía, sostenibilidad": [
            "https://www.portafolio.co/rss/energia.xml",
            "https://www.portafolio.co/rss/sostenibilidad.xml",
        ],
        "Negocios, mis finanzas": [
            "https://www.portafolio.co/rss/negocios.xml",
            "https://www.portafolio.co/rss/mis-finanzas.xml",
        ],
    }
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/rss+xml, application/xml, text/xml, */*",
    "Accept-Language": "es-CO,es;q=0.9",
}

GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
GEMINI_TIMEOUT = 25

# Color de texto por categorias (se comparten algunas, otras usan general)
CATEGORY_COLORS = {
    # El Tiempo
    "Opinion":    "4E342E",
    "Politica":   "B71C1C",
    "Economia":   "0D47A1",
    "Deportes":   "1B5E20",
    "Mundo":      "4A148C",
    "Bogota":     "E65100",
    "Colombia":   "006064",
    "Cultura":    "00695C",
    "Vida":       "33691E",
    "Justicia":   "37474F",
    "Salud":      "880E4F",
    # Portafolio (categorias agrupadas)
    "Economía, indicadores":              "0D47A1",
    "Tendencias":                          "00838F",
    "Internacional, emprendimiento, tecno": "4A148C",
    "Energía, sostenibilidad":             "D84315",
    "Negocios, mis finanzas":              "1565C0",
    # General
    "General":    "424242",
}

PREFIX_NOISE = [
    r"^en vivo\s*[:\-\—\–]",
    r"^en directo\s*[:\-\—\–]",
    r"^video\s*[:\-\—\–]",
    r"^fotos\s*[:\-\—\–]",
    r"^en fotos\s*[:\-\—\–]",
    r"^galería\s*[:\-\—\–]",
    r"^atención\s*[:\-\—\–]",
    r"^urgente\s*[:\-\—\–]",
    r"^análisis\s*[:\-\—\–]",
    r"^opinión\s*[:\-\—\–]",
    r"^ojo\s*[:\-\—\–]",
    r"^minuto a minuto\s*[:\-\—\–]",
    r"^lo último\s*[:\-\—\–]",
]

HANGING_WORDS = {
    "de", "del", "en", "para", "con", "por", "y", "e", "o", "u", "a", "al",
    "tras", "el", "la", "los", "las", "un", "una", "unos", "unas", "su", "sus",
    "sin", "sobre", "entre", "hacia", "hasta", "desde", "ante", "bajo", "contra",
    "que", "se", "es", "como", "mas", "más", "pero", "sino", "este", "esta",
    "estos", "estas", "ese", "esa", "esos", "esas", "us", "usd", "durante",
    "según", "segun", "incluyendo", "solo", "sólo", "cada", "aunque", "porque",
    "cuando", "donde", "mientras", "tan", "muy"
}

STOP_WORDS = {
    "de", "del", "en", "para", "con", "por", "y", "e", "o", "u", "a", "al",
    "tras", "el", "la", "los", "las", "un", "una", "unos", "unas", "su", "sus",
    "sin", "sobre", "entre", "hacia", "hasta", "desde", "ante", "bajo", "contra",
    "que", "se", "es", "como", "mas", "más", "pero", "sino", "este", "esta",
    "estos", "estas", "ese", "esa", "esos", "esas", "durante", "según", "segun",
    "solo", "sólo", "cada", "aunque", "porque", "cuando", "donde", "mientras", "tan", "muy"
}

# ─────────────────────────────────────────────
# UTILIDADES Y PROCESAMIENTO
# ─────────────────────────────────────────────

def clean_and_format_title(raw_title: str, max_words: int = 10, max_chars: int = 90) -> str:
    """
    Limpia y da formato profesional al titular periodístico para Flowcards:
    - Conserva total sentido gramatical, coherencia y sintaxis natural.
    - Elimina prefijos de ruido ('EXCLUSIVO:', 'EN VIVO:', 'URGENTE:', etc.).
    - Elimina comillas innecesarias y autorías al final.
    """
    if not raw_title or not str(raw_title).strip():
        return "Sin título"
    
    text = str(raw_title).strip()
    
    # 1. Quitar prefijos periodísticos comunes
    for pattern in PREFIX_NOISE:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE).strip()
        
    # 2. Quitar citas/autorías al final (ej: ', según expertos' o ': James Rockall')
    text = re.sub(r':\s*[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,3}\s*$', '', text)
    text = re.sub(r',\s*según\s+.*$', '', text, flags=re.IGNORECASE)
    
    # 3. Limpiar signos extraños al inicio o fin
    text = text.strip('"\'“”«» ').rstrip('.,;:-')
    
    # 4. Si el titular es excesivamente largo, cortar en un límite natural de palabras
    words = text.split()
    if len(words) > max_words:
        text = " ".join(words[:max_words]).rstrip('.,;:-')
        
    return text if text else str(raw_title).strip()

def shorten_title_smart(title: str, max_words: int = 10) -> str:
    """Función de compatibilidad que llama a clean_and_format_title."""
    return clean_and_format_title(title, max_words=max_words)

def parse_pubdate(pubdate_str: str):
    if not pubdate_str:
        dt = datetime.now()
        local = dt.astimezone()
        hora_12 = local.strftime("%I:%M %p").lstrip("0")
        return local.strftime("%Y-%m-%d"), hora_12, local.timestamp()
    try:
        dt = parsedate_to_datetime(pubdate_str)
    except Exception:
        try:
            dt = datetime.fromisoformat(pubdate_str)
        except Exception:
            dt = datetime.now()
            local = dt.astimezone()
            hora_12 = local.strftime("%I:%M %p").lstrip("0")
            return local.strftime("%Y-%m-%d"), hora_12, local.timestamp()
    local = dt.astimezone()
    hora_12 = local.strftime("%I:%M %p").lstrip("0")
    return local.strftime("%Y-%m-%d"), hora_12, local.timestamp()

def make_amp_url(url: str, url_base: str) -> str:
    """Inserta /amp/ en la URL específica del diario."""
    if not url:
        return url
    target = url_base + "/"
    amp_target = url_base + "/amp/"
    return url.replace(target, amp_target, 1)


def analyze_seo_potential(article: dict) -> dict:
    """Asigna un puntaje heuristico de potencial SEO/tendencia para El Tiempo."""
    titulo_raw = str(article.get("titulo_raw", "") or "")
    resumen = str(article.get("resumen", "") or "")
    categoria = str(article.get("categoria", "") or "")
    combined_text = f"{titulo_raw} {resumen}".lower()

    score = 35
    reasons = []

    age_hours = max(0.0, (datetime.now().astimezone().timestamp() - float(article.get("timestamp", 0))) / 3600)
    if age_hours <= 6:
        score += 16
        reasons.append("muy reciente")
    elif age_hours <= 12:
        score += 12
        reasons.append("reciente")
    elif age_hours <= 24:
        score += 8
        reasons.append("aun fresca")
    elif age_hours <= 48:
        score += 4

    category_bonus = CATEGORY_POTENTIAL.get(categoria, 4)
    score += category_bonus
    if category_bonus >= 10:
        reasons.append(f"tema masivo de {categoria.lower()}")
    elif category_bonus >= 7:
        reasons.append(f"categoria activa: {categoria.lower()}")

    trend_hits = sorted({kw for kw in TREND_KEYWORDS if kw in combined_text})
    if trend_hits:
        score += min(18, len(trend_hits) * 5)
        reasons.append("tema con traccion de actualidad")

    seo_hits = sorted({kw for kw in SEO_KEYWORDS if kw in combined_text})
    if seo_hits:
        score += min(14, len(seo_hits) * 4)
        reasons.append("angulo util para busqueda")

    if re.search(r"\b\d+\b", titulo_raw):
        score += 7
        reasons.append("titulo con dato concreto")

    title_words = len([w for w in re.split(r"\s+", titulo_raw.strip()) if w])
    if 7 <= title_words <= 15:
        score += 7
        reasons.append("titulo con longitud competitiva")
    elif 5 <= title_words <= 18:
        score += 4

    title_chars = len(titulo_raw)
    if 45 <= title_chars <= 110:
        score += 6
    elif 30 <= title_chars <= 130:
        score += 3

    summary_len = len(resumen)
    if 90 <= summary_len <= 260:
        score += 5
    elif summary_len >= 45:
        score += 2

    if categoria == "Opinion":
        score -= 8

    score = max(0, min(100, score))

    if score >= 78:
        level = "Muy alto"
    elif score >= 63:
        level = "Alto"
    elif score >= 48:
        level = "Medio"
    else:
        level = "Bajo"

    seo_reason = ", ".join(reasons[:3]) if reasons else "senal limitada para tendencia o SEO"
    return {
        "seo_score": score,
        "seo_level": level,
        "seo_reason": seo_reason,
    }


def load_dotenv_simple():
    """Carga variables desde el archivo .env si existe."""
    env_paths = [".env", os.path.join(os.path.dirname(__file__), ".env")]
    for p in env_paths:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#") and "=" in line:
                            k, v = line.split("=", 1)
                            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))
            except Exception:
                pass

load_dotenv_simple()

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "openai/gpt-oss-120b"
GROQ_TIMEOUT = 25

def get_groq_api_key() -> str:
    load_dotenv_simple()
    return (
        os.environ.get("GROQ_API_KEY")
        or os.environ.get("GROQ_KEY")
        or ""
    ).strip()

def analyze_with_groq(article: dict) -> dict | None:
    api_key = get_groq_api_key()
    if not api_key:
        return None

    raw_title = article.get("titulo_raw", "")
    category = article.get("categoria", "")
    summary = article.get("resumen", "")

    prompt = (
        f"Eres un editor periodístico senior de El Tiempo / Portafolio.\n"
        f"Tu tarea es generar un titular para Flowcard que sea DIRECTO, IMPACTANTE y con TOTAL SENTIDO GRAMATICAL Y PERIODÍSTICO.\n\n"
        f"DATOS DE LA NOTICIA:\n"
        f"- Categoría: {category}\n"
        f"- Titular original: {raw_title}\n"
        f"- Resumen: {summary}\n\n"
        f"REGLAS PARA EL TITULAR FLOWCARD:\n"
        f"1. DEBE tener sentido completo, redacción natural y sintaxis impecable en español.\n"
        f"2. NO recortes palabras ni dejes frases incoherentes o palabras sueltas sin conexión.\n"
        f"3. Longitud ideal: 4 a 8 palabras claras que comuniquen con fuerza lo más importante de la noticia.\n"
        f"4. Evalúa el potencial SEO y de tendencia de 0 a 100.\n\n"
        f'Responde ÚNICAMENTE en formato JSON estructurado:\n'
        f'{{\n'
        f'  "titulo_flowcard": "Titular con sentido completo y alto impacto",\n'
        f'  "seo_score": 88,\n'
        f'  "seo_level": "Muy alto",\n'
        f'  "seo_reason": "Explicación breve del interés periodístico",\n'
        f'  "keyword_objetivo": "término clave principal",\n'
        f'  "trend_type": "Nacional / Economía / Tendencia"\n'
        f'}}'
    )

    # 1. Intentar con Groq SDK oficial
    try:
        from groq import Groq
        client = Groq(api_key=api_key)
        
        models_to_try = [
            "openai/gpt-oss-120b",
            "llama-3.3-70b-versatile",
            "llama-3.1-70b-versatile",
            "llama3-70b-8192"
        ]
        
        for m in models_to_try:
            try:
                chat_completion = client.chat.completions.create(
                    model=m,
                    messages=[
                        {"role": "system", "content": "Eres un editor periodístico SEO experto. Respondes únicamente en formato JSON válido."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.4,
                    max_completion_tokens=1024,
                    response_format={"type": "json_object"}
                )
                content_resp = chat_completion.choices[0].message.content
                if content_resp:
                    data_json = json.loads(content_resp)
                    score = int(data_json.get("seo_score", 80))
                    score = max(0, min(100, score))
                    titulo_flow = str(data_json.get("titulo_flowcard", "")).strip()
                    if not titulo_flow or len(titulo_flow.split()) < 2:
                        titulo_flow = clean_and_format_title(raw_title)
                    
                    return {
                        "seo_score": score,
                        "seo_level": str(data_json.get("seo_level", "Alto")).strip(),
                        "seo_reason": str(data_json.get("seo_reason", ""))[:140] or "Análisis SEO con IA",
                        "keyword_objetivo": str(data_json.get("keyword_objetivo", ""))[:60],
                        "trend_type": str(data_json.get("trend_type", ""))[:40],
                        "titulo_flowcard": titulo_flow,
                        "analysis_source": f"Groq ({m})",
                    }
            except Exception:
                continue
    except Exception:
        pass

    # 2. Fallback vía requests
    try:
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        for m in ["openai/gpt-oss-120b", "llama-3.3-70b-versatile", "llama3-70b-8192"]:
            try:
                payload = {
                    "model": m,
                    "messages": [
                        {"role": "system", "content": "Eres un editor periodístico SEO experto. Respondes exclusivamente en formato JSON estructurado."},
                        {"role": "user", "content": prompt}
                    ],
                    "temperature": 0.4,
                    "response_format": {"type": "json_object"}
                }
                r = requests.post(GROQ_API_URL, headers=headers, json=payload, timeout=GROQ_TIMEOUT)
                if r.status_code == 200:
                    data = r.json()
                    text = data.get("choices", [{}])[0].get("message", {}).get("content", "")
                    if text:
                        data_json = json.loads(text)
                        score = int(data_json.get("seo_score", 80))
                        score = max(0, min(100, score))
                        titulo_flow = str(data_json.get("titulo_flowcard", "")).strip()
                        if not titulo_flow or len(titulo_flow.split()) < 2:
                            titulo_flow = clean_and_format_title(raw_title)
                        return {
                            "seo_score": score,
                            "seo_level": str(data_json.get("seo_level", "Alto")).strip(),
                            "seo_reason": str(data_json.get("seo_reason", ""))[:140] or "Análisis SEO con IA",
                            "keyword_objetivo": str(data_json.get("keyword_objetivo", ""))[:60],
                            "trend_type": str(data_json.get("trend_type", ""))[:40],
                            "titulo_flowcard": titulo_flow,
                            "analysis_source": f"Groq ({m})",
                        }
            except Exception:
                continue
    except Exception as exc:
        print(f"   [IA Groq] Fallback heurístico: {exc}")

    return None

def shorten_to_25_chars(title: str) -> str:
    """Asegura que un titular tenga máximo 25 caracteres sin cortar palabras torpemente."""
    if not title:
        return "Sin titulo"
    clean = re.sub(r'[^\w\s]', '', str(title).strip(), flags=re.UNICODE)
    if len(clean) <= 25:
        return clean
    words = clean.split()
    result = []
    current_len = 0
    for w in words:
        projected = current_len + len(w) + (1 if result else 0)
        if projected > 25:
            break
        result.append(w)
        current_len = projected
    if not result:
        return clean[:25]
    return " ".join(result)


def generate_temas_del_dia_headline(article: dict) -> str:
    """Genera un titular para 'Temas del Día' de máximo 25 caracteres con IA Groq y sentido completo."""
    api_key = get_groq_api_key()
    raw_title = article.get("titulo_raw", "")
    summary = article.get("resumen", "")
    
    if api_key:
        prompt = (
            f"Eres un editor de portadas. Crea un titular periodístico de impacto para 'Temas del Día'.\n"
            f"REGLA ESTRICTA:\n"
            f"- El titular DEBE tener MÁXIMO 25 CARACTERES TOTALES (incluyendo letras y espacios).\n"
            f"- Debe tener SENTIDO COMPLETO, palabras reales y gramática correcta.\n"
            f"- Ejemplos excelentes: 'Petro anuncia reformas' (22), 'Dólar baja a mínimo' (19), 'Alerta roja por lluvias' (22).\n\n"
            f"Noticia:\n"
            f"Título original: {raw_title}\n"
            f"Resumen: {summary}\n\n"
            f'Responde ÚNICAMENTE en formato JSON: {{"titular_25": "Titular <= 25 chars"}}'
        )
        try:
            from groq import Groq
            client = Groq(api_key=api_key)
            for m in ["openai/gpt-oss-120b", "llama-3.3-70b-versatile", "llama3-70b-8192"]:
                try:
                    resp = client.chat.completions.create(
                        model=m,
                        messages=[
                            {"role": "system", "content": "Eres un editor de portadas breves. Creas titulares con límite estricto de 25 caracteres y sentido completo."},
                            {"role": "user", "content": prompt}
                        ],
                        temperature=0.3,
                        max_completion_tokens=256,
                        response_format={"type": "json_object"}
                    )
                    content = resp.choices[0].message.content
                    if content:
                        parsed = json.loads(content)
                        t25 = str(parsed.get("titular_25", "")).strip()
                        if t25 and len(t25) <= 25:
                            return t25
                        elif t25:
                            return shorten_to_25_chars(t25)
                except Exception:
                    continue
        except Exception:
            pass

    return shorten_to_25_chars(raw_title)

def export_temas_del_dia(articles: list, work_dir: str) -> str | None:
    """
    Selecciona las mejores 8 noticias de El Tiempo, genera para cada una un titular de <=25 caracteres con IA Groq
    y actualiza la plantilla 'TEMAS DEL DÍA.xlsx'.
    """
    if not articles:
        print("\n[!] No hay artículos para generar Temas del Día.")
        return None

    template_file = os.path.join(work_dir, "TEMAS DEL DÍA.xlsx")
    if not os.path.exists(template_file):
        template_file = "TEMAS DEL DÍA.xlsx"
        if not os.path.exists(template_file):
            print(f"\n[!] No se encontró la plantilla 'TEMAS DEL DÍA.xlsx'")
            return None

    # Ordenar noticias por puntaje SEO o frescura y elegir las 8 mejores
    top_8 = sorted(
        articles,
        key=lambda x: (x.get("seo_score", 0), x.get("timestamp", 0)),
        reverse=True
    )[:8]

    print(f"\n--- PROCESANDO TEMAS DEL DÍA (Top 8 noticias con IA Groq) ---")

    try:
        wb = openpyxl.load_workbook(template_file)
        ws = wb.active

        for idx, art in enumerate(top_8, 1):
            row = idx + 2  # Filas 3 a 10
            titular_25 = generate_temas_del_dia_headline(art)
            url_target = art.get("url_amp") or art.get("url_original", "")

            ws.cell(row=row, column=1, value=idx)                # Col A: RANKING
            ws.cell(row=row, column=2, value=f"=+LEN(C{row})")  # Col B: CARACTERES
            ws.cell(row=row, column=3, value=titular_25)         # Col C: TÍTULO (<= 25 chars)
            
            cell_url = ws.cell(row=row, column=4, value=url_target) # Col D: URL
            if url_target:
                cell_url.hyperlink = url_target
                cell_url.font = Font(name="Century Gothic", color="0277BD", underline="single", size=10)

            print(f"   Top {idx}: [{len(titular_25)} chars] \"{titular_25}\" -> {url_target[:50]}...")

        # Si hay menos de 8 noticias, limpiar filas sobrantes
        for idx in range(len(top_8) + 1, 9):
            row = idx + 2
            ws.cell(row=row, column=1, value=idx)
            ws.cell(row=row, column=2, value=f"=+LEN(C{row})")
            ws.cell(row=row, column=3, value="")
            ws.cell(row=row, column=4, value="")

        wb.save(template_file)
        print(f"\n[LISTO] Guardado 'TEMAS DEL DÍA.xlsx' con el Top 8 de noticias.")
        return template_file
    except PermissionError:
        print("\n[ERROR CRÍTICO] 'TEMAS DEL DÍA.xlsx' está abierto en Excel. Ciérrelo para actualizar.")
        return None
    except Exception as e:
        print(f"\n[ERROR] Falló la exportación de Temas del Día: {e}")
        return None


def is_subscriber_only(article_url: str) -> bool:
    """Verifica si un articulo de El Tiempo es exclusivo para suscriptores.
    Aplica filtros precisos para no confundir con menús JS globales."""
    try:
        resp = requests.get(article_url, headers=HEADERS, timeout=10)
        resp.raise_for_status()
        html_content = resp.text
        lower_html = html_content.lower()
        
        # 1. Coincidencias exactas en todo el HTML (para meta tags y variables JS)
        EXACT_MATCHES = [
            'content="csuscriptor-modal"',
            'content="exclusivo suscriptores"',
            '"premium": true', '"premium":true',
            'isrestricted":true'
        ]
        for match in EXACT_MATCHES:
            if match in lower_html:
                return True
                
        # 2. Análisis estructural detallado (solo dentro del contenido del artículo)
        if BeautifulSoup is not None:
            soup = BeautifulSoup(html_content, "html.parser")
            article_tag = soup.find("article")
            
            if article_tag:
                # Buscar un badge explícito de premium, pero EVITANDO las recomendaciones 
                # (elementos asides o cajas en el footer que traen c-suscriptor)
                header = article_tag.find("header") or article_tag
                
                # a. Clases CSS típicas del badge de suscriptor en el encabezado
                SUBSCRIBER_CLASSES = [
                    "badge-premium", "badge-subscriber", "premium-badge", "paywall"
                ]
                for css_class in SUBSCRIBER_CLASSES:
                    if header.find(class_=lambda c: c and css_class in c.lower()):
                        return True
                        
                # b. Atributos data- de paywall en el html local
                PAYWALL_ATTRS = ["data-premium", "data-paywall"]
                for attr in PAYWALL_ATTRS:
                    if header.find(attrs={attr: True}):
                        return True
                        
                # c. Texto ET visible en un badge pequeño (exclusivo ET) en el header
                for tag in header.find_all(["span", "div", "label"]):
                    # Verificar si la clase del tag sugiere un badge
                    tag_class = str(tag.get("class", "")).lower()
                    if tag.get_text(strip=True) == "ET" and ("badge" in tag_class or "label" in tag_class):
                        return True

    except Exception:
        pass  # Si falla, asumimos que NO es exclusivo
    return False


# ─────────────────────────────────────────────
# SCRAPING
# ─────────────────────────────────────────────

def fetch_top_article(feed_url: str, label: str, url_base: str, seen_urls: set, check_paywall: bool = False) -> dict | None:
    """Extrae el primer articulo valido de un feed RSS (el mas reciente).
    Si check_paywall=True, verifica que no sea exclusivo para suscriptores."""
    print(f"\n>> [{label}] Buscando en {feed_url}")
    try:
        resp = requests.get(feed_url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"   [!] Omitido (No accesible o no existe el feed).")
        return None

    try:
        root = ET.fromstring(resp.content)
    except ET.ParseError as exc:
        print(f"   [!] XML invalido.")
        return None

    items = root.findall(".//item")
    if not items:
        print(f"   [!] Feed vacío.")
        return None

    def _get_text(element) -> str:
        if element is not None and isinstance(element.text, str):
            val: str = element.text
            return val.strip()
        return ""

    parsed_items = []
    
    for item in items:
        link_el = item.find("link")
        art_url = _get_text(link_el)
        if not art_url:
            guid_el = item.find("guid")
            art_url = _get_text(guid_el)
            
        if not art_url:
            continue
            
        title_el = item.find("title")
        titulo_raw = _get_text(title_el)

        pub_el = item.find("pubDate")
        pub_str = _get_text(pub_el)
        fecha, hora, timestamp = parse_pubdate(pub_str)

        desc_el = item.find("description")
        resumen = _get_text(desc_el)

        cat_el = item.find("category")
        cat_raw = _get_text(cat_el)
        
        parsed_items.append({
            "art_url": art_url,
            "titulo_raw": titulo_raw,
            "fecha": fecha,
            "hora": hora,
            "timestamp": timestamp,
            "resumen": resumen,
            "cat_raw": cat_raw
        })

    # Ordenar chronológicamente descendente para asegurar que el primero es el más reciente
    parsed_items.sort(key=lambda x: x["timestamp"], reverse=True)

    for item_data in parsed_items:
        art_url = item_data["art_url"]

        if art_url in seen_urls:
            continue

        # Si es Portafolio, filtrar y omitir cualquier nota de la categoría u origen Opinión
        if "portafolio" in url_base.lower():
            cat_check = str(item_data.get("cat_raw", "") or "").lower()
            url_check = str(art_url or "").lower()
            if "opinion" in cat_check or "opinión" in cat_check or "/opinion/" in url_check:
                seen_urls.add(art_url)
                print(f"   [!] Omitido (Categoría Opinión en Portafolio): {item_data['titulo_raw'][:60]}")
                continue

        if check_paywall and is_subscriber_only(art_url):
            file_name_part = str(art_url.split('/')[-1])
            print(f"   [!] Omitido (Solo para suscriptores): {file_name_part[:60]}")
            seen_urls.add(art_url)
            continue

        seen_urls.add(art_url)
        
        # Mapeo por defecto o nombre del label
        categoria = label
        if label == "Opinion":
            cat_map = {"Columnistas": "Opinion", "Editorial": "Opinion", "Cartas": "Opinion", "Caricaturas": "Opinion"}
            categoria = cat_map.get(item_data["cat_raw"], label)
        titulo_3 = clean_and_format_title(item_data["titulo_raw"])
        url_amp  = make_amp_url(art_url, url_base)

        article_data = {
            "titulo_3":     titulo_3,
            "titulo_raw":   item_data["titulo_raw"],
            "url_amp":      url_amp,
            "url_original": art_url,
            "categoria":    categoria,
            "fecha":        item_data["fecha"],
            "hora":         item_data["hora"],
            "timestamp":    item_data["timestamp"],
            "resumen":      item_data["resumen"],
        }
        if "eltiempo" in url_base.lower():
            ai_analysis = analyze_with_groq(article_data)
            if ai_analysis:
                if ai_analysis.get("titulo_flowcard"):
                    article_data["titulo_3"] = ai_analysis["titulo_flowcard"]
                    titulo_3 = article_data["titulo_3"]
                article_data.update(ai_analysis)
            else:
                heuristic_analysis = analyze_seo_potential(article_data)
                heuristic_analysis["analysis_source"] = "Heuristico"
                article_data.update(heuristic_analysis)

        print(f"   OK -> \"{titulo_3}\" | {item_data['fecha']} {item_data['hora']}")
        return article_data

    print(f"   [!] Omitido (Todos los artículos de esta categoría ya fueron agregados o descartados).")
    return None


def fetch_top_from_group(feed_urls: list, label: str, url_base: str, seen_urls: set, check_paywall: bool = False) -> dict | None:
    """Busca en multiples feeds RSS y devuelve el articulo mas reciente no repetido.
    Usado para las categorias agrupadas de Portafolio.
    
    NOTA: Para evitar que los candidatos descartados bloqueen futuros grupos,
    se usa un conjunto temporal y solo el ganador se agrega al seen_urls global."""
    candidates = []
    candidate_urls = []  # URLs de los candidatos para revertir los no-ganadores
    
    for feed_url in feed_urls:
        # Usamos seen_urls normal: ya cubre URLs ya confirmadas en grupos anteriores
        art = fetch_top_article(feed_url, label, url_base, seen_urls, check_paywall=check_paywall)
        if art:
            candidates.append(art)
            candidate_urls.append(art["url_original"])
    
    if not candidates:
        return None
    
    # Ordenar por fecha+hora descendente y devolver el mas reciente
    candidates.sort(key=lambda x: x.get("timestamp", 0), reverse=True)
    winner = candidates[0]
    
    # Revertir del seen_urls los candidatos que NO ganaron,
    # para que otros grupos puedan considerar esos artículos si los necesitan.
    for art in candidates[1:]:
        seen_urls.discard(art["url_original"])
    
    return winner


# ─────────────────────────────────────────────
# EXCEL - DISEÑO Y EXPORTACIÓN
# ─────────────────────────────────────────────

def thin_border():
    s = Side(style="thin", color="B0BEC5")
    return Border(left=s, right=s, top=s, bottom=s)

def export_to_excel(articles: list, output_path: str, book_name: str, include_amp: bool = True) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = book_name
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 25

    is_el_tiempo = book_name == "El Tiempo"
    if include_amp:
        encabezados = ["✔", "Titulo", "URL SIN AMP", "CATEGORÍA", "Fecha", "URL AMP", "Resumen"]
    else:
        encabezados = ["✔", "Titulo", "URL SIN AMP", "CATEGORÍA", "Fecha", "Resumen"]
    if is_el_tiempo:
        encabezados.extend(["POTENCIAL", "PUNTAJE SEO", "MOTIVO SEO", "KEYWORD", "FUENTE ANALISIS"])

    hdr_fill = PatternFill("solid", fgColor="000000")
    hdr_font = Font(name="Century Gothic", bold=True, color="FFFFFF", size=11)
    hdr_align = Alignment(horizontal="center", vertical="center")

    for col, texto in enumerate(encabezados, 1):
        cell = ws.cell(row=1, column=col, value=texto)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = hdr_align
        cell.border = thin_border()

    last_column = get_column_letter(len(encabezados))
    ws.auto_filter.ref = f"A1:{last_column}1"

    fill_a = PatternFill("solid", fgColor="EEF2FA")
    fill_b = PatternFill("solid", fgColor="FFFFFF")

    if is_el_tiempo:
        sorted_arts = sorted(
            articles,
            key=lambda x: (x.get("seo_score", 0), x.get("timestamp", 0)),
            reverse=True,
        )
    else:
        sorted_arts = sorted(articles, key=lambda x: x.get("timestamp", 0), reverse=True)
    last_row = len(sorted_arts) + 1

    for r, art in enumerate(sorted_arts, 2):
        relleno = fill_a if r % 2 == 0 else fill_b
        fecha_hora = f"{art['fecha']}  {art['hora']}"

        # col_offset: cuando no hay AMP, las columnas a partir de Resumen se desplazan -1
        amp_shift = 0 if include_amp else -1

        ca = ws.cell(row=r, column=1, value="")
        ca.border = thin_border()
        ca.fill = relleno
        ca.font = Font(name="Century Gothic", bold=True, size=14, color="1B5E20")
        ca.alignment = Alignment(horizontal="center", vertical="center")

        cb = ws.cell(row=r, column=2, value=art["titulo_3"])
        cb.border = thin_border()
        cb.fill = relleno
        cb.font = Font(name="Century Gothic", bold=True, size=11)
        cb.alignment = Alignment(horizontal="left", vertical="center")

        cc = ws.cell(row=r, column=3, value=art["url_original"])
        cc.border = thin_border()
        cc.fill = relleno
        cc.alignment = Alignment(horizontal="left", vertical="center")
        if art["url_original"]:
            cc.hyperlink = art["url_original"]
            cc.font = Font(name="Century Gothic", color="0277BD", underline="single", size=10)
        else:
            cc.font = Font(name="Century Gothic", size=10)

        cat_color = CATEGORY_COLORS.get(art["categoria"], "424242")
        cd = ws.cell(row=r, column=4, value=art["categoria"])
        cd.border = thin_border()
        cd.fill = relleno
        cd.font = Font(name="Century Gothic", bold=True, color=cat_color, size=11)
        cd.alignment = Alignment(horizontal="center", vertical="center")

        ce = ws.cell(row=r, column=5, value=fecha_hora)
        ce.border = thin_border()
        ce.fill = relleno
        ce.font = Font(name="Century Gothic", size=10)
        ce.alignment = Alignment(horizontal="center", vertical="center")

        # Columna URL AMP (solo si include_amp=True)
        if include_amp:
            cf = ws.cell(row=r, column=6, value=art["url_amp"])
            cf.border = thin_border()
            cf.fill = relleno
            cf.alignment = Alignment(horizontal="left", vertical="center")
            if art["url_amp"]:
                cf.hyperlink = art["url_amp"]
                cf.font = Font(name="Century Gothic", color="1565C0", underline="single", size=10)
            else:
                cf.font = Font(name="Century Gothic", size=10)

        # Columna Resumen: col 7 con AMP, col 6 sin AMP
        cg = ws.cell(row=r, column=7 + amp_shift, value=art.get("resumen", ""))
        cg.border = thin_border()
        cg.fill = relleno
        cg.font = Font(name="Century Gothic", size=9, color="424242")
        cg.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        if is_el_tiempo:
            seo_level = art.get("seo_level", "N/D")
            seo_score = art.get("seo_score", 0)
            seo_reason = art.get("seo_reason", "")

            level_color = "2E7D32"
            if seo_level == "Alto":
                level_color = "1565C0"
            elif seo_level == "Medio":
                level_color = "EF6C00"
            elif seo_level == "Bajo":
                level_color = "6D4C41"

            ch = ws.cell(row=r, column=8 + amp_shift, value=seo_level)
            ch.border = thin_border()
            ch.fill = relleno
            ch.font = Font(name="Century Gothic", bold=True, size=10, color=level_color)
            ch.alignment = Alignment(horizontal="center", vertical="center")

            ci = ws.cell(row=r, column=9 + amp_shift, value=seo_score)
            ci.border = thin_border()
            ci.fill = relleno
            ci.font = Font(name="Century Gothic", bold=True, size=10)
            ci.alignment = Alignment(horizontal="center", vertical="center")

            cj = ws.cell(row=r, column=10 + amp_shift, value=seo_reason)
            cj.border = thin_border()
            cj.fill = relleno
            cj.font = Font(name="Century Gothic", size=9, color="424242")
            cj.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            ck = ws.cell(row=r, column=11 + amp_shift, value=art.get("keyword_objetivo", ""))
            ck.border = thin_border()
            ck.fill = relleno
            ck.font = Font(name="Century Gothic", size=9, color="424242")
            ck.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            cl = ws.cell(row=r, column=12 + amp_shift, value=art.get("analysis_source", "Heuristico"))
            cl.border = thin_border()
            cl.fill = relleno
            cl.font = Font(name="Century Gothic", bold=True, size=9, color="424242")
            cl.alignment = Alignment(horizontal="center", vertical="center")

    dv = DataValidation(
        type="list",
        formula1='" ,✔"',
        allow_blank=True,
        showDropDown=False,
    )
    dv.prompt = "Escribe ✔ para marcar"
    dv.promptTitle = "Marcar artículo"
    dv.error = "Solo se permite el signo: ✔ o celda vacía"
    dv.errorTitle = "Valor no válido"

    if last_row >= 2:
        dv.add(f"A2:A{last_row}")
        ws.add_data_validation(dv)

        green_fill = PatternFill("solid", fgColor="C8E6C9")
        green_font_bold = Font(name="Century Gothic", bold=True, color="1B5E20", size=11)
        rule = FormulaRule(
            formula=['LEN(TRIM($A2))>0'],
            fill=green_fill,
            font=green_font_bold,
        )
        ws.conditional_formatting.add(f"A2:{last_column}{last_row}", rule)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 22
    if include_amp:
        ws.column_dimensions["F"].width = 65   # URL AMP
        ws.column_dimensions["G"].width = 55   # Resumen
        if is_el_tiempo:
            ws.column_dimensions["H"].width = 16
            ws.column_dimensions["I"].width = 14
            ws.column_dimensions["J"].width = 42
            ws.column_dimensions["K"].width = 24
            ws.column_dimensions["L"].width = 18
    else:
        ws.column_dimensions["F"].width = 55   # Resumen (desplazado)
        if is_el_tiempo:
            ws.column_dimensions["G"].width = 16
            ws.column_dimensions["H"].width = 14
            ws.column_dimensions["I"].width = 42
            ws.column_dimensions["J"].width = 24
            ws.column_dimensions["K"].width = 18

    try:
        wb.save(output_path)
        print(f"\n[LISTO] Guardado con éxito el libro: {output_path}")
    except PermissionError:
        print(f"\n[ERROR CRÍTICO] Microsoft Excel tiene abierto el archivo '{book_name}.xlsx'.")
        print("POR FAVOR CIERRE EL EJECUTABLE DE EXCEL O EL ARCHIVO E INICIE LA BÚSQUEDA NUEVAMENTE.")
        raise RuntimeError(f"El archivo {book_name}.xlsx está abierto o bloqueado. Ciérrelo primero.")


# ─────────────────────────────────────────────
# BLOQUE PRINCIPAL
# ─────────────────────────────────────────────

def run_scraper():
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    print("=" * 60)
    print("  AUTOMATIZACIÓN SANTAMARÍA — Generador de Flowcards")
    print(f"  Fecha: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}")
    print("=" * 60)

    # Cálculo de total de tareas para el porcentaje
    total_tasks = sum(len(config["feeds"]) for config in SITES.values())
    total_tasks += len(PORTAFOLIO_CONFIG.get("feeds_grouped", {}))
    current_task = 0

    # Directorio actual
    work_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.exists(work_dir):
        work_dir = "./"

    history_file = os.path.join(work_dir, "historial_urls.json")
    global_seen_urls = set()
    if os.path.exists(history_file):
        try:
            with open(history_file, "r", encoding="utf-8") as f:
                global_seen_urls = set(json.load(f))
        except Exception:
            pass

    # ── El Tiempo (feeds simples, con filtro de suscriptores) ──
    for site_name, config in SITES.items():
        print(f"\n\n--- INICIANDO EXTRACCIÓN PARA: {site_name} ---")
        out_path = os.path.join(work_dir, config["output_file"])
        url_base = config["url_base"]
        is_eltiempo = ("eltiempo" in url_base)
        
        articles = []
        # Conjunto de deduplicacion propio para El Tiempo (copia del historial global)
        seen_urls_et = set(global_seen_urls)
        seen_titles_et = set()  # Dedup por titulo para El Tiempo
        for cat_label, url in config["feeds"].items():
            pct = int((current_task / total_tasks) * 100)
            print(f"\nEspere un momento, estamos buscando noticias para sus Flowcards. ({pct}%)")
            art = fetch_top_article(url, cat_label, url_base, seen_urls_et, check_paywall=is_eltiempo)
            current_task += 1
            if art:
                title_key = art["titulo_raw"].strip().lower()
                if title_key not in seen_titles_et:
                    seen_titles_et.add(title_key)
                    articles.append(art)
                else:
                    print(f"   [!] Omitido (titulo duplicado entre categorias): {art['titulo_raw'][:60]}")
        # Agregar las URLs nuevas de El Tiempo al historial global
        global_seen_urls.update(seen_urls_et)
                
        if not articles:
            print(f"\n[!] No se encontraron articulos aptos para {site_name}.")
            continue
            
        export_to_excel(articles, out_path, site_name)

    # ── Portafolio (feeds agrupados, con filtro de suscriptores) ──
    print(f"\n\n--- INICIANDO EXTRACCIÓN PARA: Portafolio ---")
    p_config = PORTAFOLIO_CONFIG
    p_out_path = os.path.join(work_dir, str(p_config["output_file"]))
    p_url_base = str(p_config["url_base"])
    is_portafolio = ("portafolio" in p_url_base.lower())
    
    p_articles = []
    # Conjunto de deduplicacion PROPIO para Portafolio (independiente de El Tiempo)
    p_seen_urls = set(global_seen_urls)
    p_seen_titles = set()  # Dedup por titulo entre grupos de Portafolio
    feeds_grouped_dict: dict = p_config.get("feeds_grouped", {})  # type: ignore
    for group_label, feed_list in feeds_grouped_dict.items():
        pct = int((current_task / total_tasks) * 100)
        print(f"\nEspere un momento, estamos buscando noticias para sus Flowcards. ({pct}%)")
        art = fetch_top_from_group(feed_list, group_label, p_url_base, p_seen_urls, check_paywall=is_portafolio)
        current_task += 1
        if art:
            title_key = art["titulo_raw"].strip().lower()
            if title_key not in p_seen_titles:
                p_seen_titles.add(title_key)
                p_articles.append(art)
            else:
                print(f"   [!] Omitido (titulo duplicado entre grupos Portafolio): {art['titulo_raw'][:60]}")
    # Agregar las URLs nuevas de Portafolio al historial global
    global_seen_urls.update(p_seen_urls)
    
    # 100% al finalizar las búsquedas
    print(f"\nEspere un momento, estamos buscando noticias para sus Flowcards. (100%)")

    if p_articles:
        export_to_excel(p_articles, p_out_path, "Portafolio")
    else:
        print(f"\n[!] No se encontraron articulos aptos para Portafolio.")

    try:
        with open(history_file, "w", encoding="utf-8") as f:
            json.dump(list(global_seen_urls), f, indent=4)
    except Exception as e:
        print(f"\n[!] No se pudo guardar el historial: {e}")

def run_scraper_selected(selected_sources=None, process_type: str = "both", include_amp: bool = True):
    if not selected_sources:
        selected_sources = ["El Tiempo", "Portafolio"]

    run_flowcards = process_type in ("flowcards", "both")
    run_temas = process_type in ("temas_del_dia", "both")

    run_el_tiempo = "El Tiempo" in selected_sources
    run_portafolio = "Portafolio" in selected_sources

    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    print("=" * 60)
    print("  FLOWCATS v2.0 - Generador de Noticias & Radar Editorial")
    print(f"  Fecha: {datetime.now().strftime('%Y-%m-%d %I:%M %p')}")
    print(f"  Modo de Proceso: {process_type.upper()} | Fuentes: {', '.join(selected_sources)}")
    print("=" * 60)

    total_tasks = 0
    if run_el_tiempo:
        total_tasks += sum(len(config["feeds"]) for config in SITES.values())
    if run_portafolio:
        total_tasks += len(PORTAFOLIO_CONFIG.get("feeds_grouped", {}))
    if total_tasks == 0:
        print(f"\n[!] No se seleccionó ningún medio para procesar.")
        return {"El Tiempo": [], "Portafolio": [], "top": []}

    current_task = 0
    work_dir = os.path.dirname(os.path.abspath(__file__))
    if not os.path.exists(work_dir):
        work_dir = "./"

    history_file = os.path.join(work_dir, "historial_urls.json")
    global_seen_urls = set()
    if os.path.exists(history_file):
        try:
            with open(history_file, "r", encoding="utf-8") as f:
                global_seen_urls = set(json.load(f))
        except Exception:
            pass

    articles = []
    if run_el_tiempo:
        for site_name, config in SITES.items():
            print(f"\n\n--- INICIANDO EXTRACCIÓN PARA: {site_name} ---")
            out_path = os.path.join(work_dir, config["output_file"])
            url_base = config["url_base"]
            is_eltiempo = ("eltiempo" in url_base)

            seen_urls_et = set(global_seen_urls)
            seen_titles_et = set()
            for cat_label, url in config["feeds"].items():
                pct = int((current_task / total_tasks) * 100)
                print(f"\nEspere un momento, buscando noticias para {site_name} - {cat_label} ({pct}%)")
                art = fetch_top_article(url, cat_label, url_base, seen_urls_et, check_paywall=is_eltiempo)
                current_task += 1
                if art:
                    title_key = art["titulo_raw"].strip().lower()
                    if title_key not in seen_titles_et:
                        seen_titles_et.add(title_key)
                        articles.append(art)
                    else:
                        print(f"   [!] Omitido (duplicado): {art['titulo_raw'][:60]}")
            global_seen_urls.update(seen_urls_et)

            if articles:
                if run_flowcards:
                    export_to_excel(articles, out_path, site_name, include_amp=include_amp)
                    print(f"\n[+] Archivo {site_name}.xlsx generado exitosamente con {len(articles)} Flowcards.")
                
                if run_temas:
                    export_temas_del_dia(articles, work_dir)
            else:
                print(f"\n[!] No se encontraron artículos aptos para {site_name}.")

    p_articles = []
    if run_portafolio:
        print(f"\n\n--- INICIANDO EXTRACCIÓN PARA: Portafolio ---")
        p_config = PORTAFOLIO_CONFIG
        p_out_path = os.path.join(work_dir, str(p_config["output_file"]))
        p_url_base = str(p_config["url_base"])
        is_portafolio = ("portafolio" in p_url_base.lower())

        p_seen_urls = set(global_seen_urls)
        p_seen_titles = set()
        feeds_grouped_dict: dict = p_config.get("feeds_grouped", {})
        for group_label, feed_list in feeds_grouped_dict.items():
            pct = int((current_task / total_tasks) * 100)
            print(f"\nEspere un momento, buscando noticias para Portafolio - {group_label} ({pct}%)")
            art = fetch_top_from_group(feed_list, group_label, p_url_base, p_seen_urls, check_paywall=is_portafolio)
            current_task += 1
            if art:
                title_key = art["titulo_raw"].strip().lower()
                if title_key not in p_seen_titles:
                    p_seen_titles.add(title_key)
                    p_articles.append(art)
                else:
                    print(f"   [!] Omitido (duplicado Portafolio): {art['titulo_raw'][:60]}")
        global_seen_urls.update(p_seen_urls)

        if p_articles:
            if run_flowcards:
                export_to_excel(p_articles, p_out_path, "Portafolio", include_amp=include_amp)
                print(f"\n[+] Archivo Portafolio.xlsx generado exitosamente con {len(p_articles)} Flowcards.")
        else:
            print(f"\n[!] No se encontraron artículos aptos para Portafolio.")

    print(f"\nEspere un momento, finalizando proceso. (100%)")

    try:
        with open(history_file, "w", encoding="utf-8") as f:
            json.dump(list(global_seen_urls), f, indent=4)
    except Exception as e:
        print(f"\n[!] No se pudo guardar el historial: {e}")

    top_8 = []
    if articles and run_temas:
        top_8 = sorted(
            articles,
            key=lambda x: (x.get("seo_score", 0), x.get("timestamp", 0)),
            reverse=True
        )[:8]

    return {
        "El Tiempo": articles if run_el_tiempo else [],
        "Portafolio": p_articles if run_portafolio else [],
        "top": top_8
    }

import threading
import os

try:
    import customtkinter as ctk
    from PIL import Image
    HAS_GUI = True
except (ImportError, ModuleNotFoundError):
    ctk = None
    Image = None
    HAS_GUI = False

class RealtimeLog:
    def __init__(self, tbox, progress_callback):
        self.tbox = tbox
        self.progress_callback = progress_callback
        
    def write(self, message):
        if message:
            self.tbox.after(0, self._insert_text, str(message))
            import re
            match = re.search(r'\(([0-9]+)%\)', str(message))
            if match:
                try:
                    pct = int(match.group(1))
                    if self.progress_callback:
                        self.tbox.after(0, self.progress_callback, pct)
                except Exception:
                    pass

    def _insert_text(self, message):
        self.tbox.configure(state="normal")
        self.tbox.insert("end", message)
        self.tbox.see("end")
        self.tbox.configure(state="disabled")

    def flush(self):
        pass

_BaseAppClass = ctk.CTk if ctk is not None else object

class FlowcatsApp(_BaseAppClass):
    def __init__(self):
        if not HAS_GUI or ctk is None:
            raise RuntimeError("CustomTkinter / Tkinter no está disponible.")
        super().__init__()
        
        # Configuracion de ventana
        self.title("Flowcats - Generador de Noticias")
        self.geometry("850x650")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        # Header Frame
        self.header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        self.header_frame.grid_columnconfigure(1, weight=1)
        
        # Logo
        try:
            import sys
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS # type: ignore
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            logo_path = os.path.join(base_path, "logo_flowcats.png")
            logo_img = ctk.CTkImage(light_image=Image.open(logo_path),
                                    dark_image=Image.open(logo_path),
                                    size=(80, 80))
            self.logo_label = ctk.CTkLabel(self.header_frame, image=logo_img, text="")
            self.logo_label.grid(row=0, column=0, padx=(0, 15))
        except Exception:
            self.logo_label = ctk.CTkLabel(self.header_frame, text="🐈‍⬛", font=("Century Gothic", 60))
            self.logo_label.grid(row=0, column=0, padx=(0, 15))

        # Titulo Textos
        self.title_label = ctk.CTkLabel(self.header_frame, text="Flowcats", font=("Century Gothic", 36, "bold"), text_color="#E0E0E0")
        self.title_label.grid(row=0, column=1, sticky="w")
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="Generador de noticias El Tiempo & Portafolio", font=("Century Gothic", 14), text_color="#A0A0A0")
        self.subtitle_label.grid(row=0, column=1, sticky="sw", pady=(40, 0))

        # Controles
        self.controls_frame = ctk.CTkFrame(self, fg_color="#1a1a1a")
        self.controls_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.controls_frame.grid_columnconfigure(3, weight=1)

        self.analysis_label = ctk.CTkLabel(
            self.controls_frame,
            text="Análisis El Tiempo: verificando...",
            font=("Century Gothic", 12, "bold"),
            text_color="#8ecae6"
        )
        self.analysis_label.grid(row=0, column=0, padx=15, pady=(12, 0), sticky="w")

        self.el_tiempo_var = ctk.BooleanVar(value=True)
        self.portafolio_var = ctk.BooleanVar(value=True)

        self.el_tiempo_check = ctk.CTkCheckBox(
            self.controls_frame,
            text="El Tiempo",
            variable=self.el_tiempo_var,
            font=("Century Gothic", 13),
            command=self.on_source_toggle,
        )
        self.el_tiempo_check.grid(row=0, column=1, padx=(0, 10), pady=(10, 0), sticky="w")

        self.portafolio_check = ctk.CTkCheckBox(
            self.controls_frame,
            text="Portafolio",
            variable=self.portafolio_var,
            font=("Century Gothic", 13),
            command=self.on_source_toggle,
        )
        self.portafolio_check.grid(row=0, column=2, padx=(0, 15), pady=(10, 0), sticky="w")

        # Separador visual
        self.amp_separator = ctk.CTkLabel(
            self.controls_frame,
            text="│",
            font=("Century Gothic", 18),
            text_color="#444444"
        )
        self.amp_separator.grid(row=0, column=3, padx=(5, 5), pady=(10, 0))

        # Checkbox AMP
        self.amp_var = ctk.BooleanVar(value=True)
        self.amp_check = ctk.CTkCheckBox(
            self.controls_frame,
            text="Incluir URL AMP",
            variable=self.amp_var,
            font=("Century Gothic", 13),
            fg_color="#1565C0",
            hover_color="#0D47A1",
        )
        self.amp_check.grid(row=0, column=4, padx=(0, 15), pady=(10, 0), sticky="w")
        
        self.btn_run = ctk.CTkButton(self.controls_frame, text="Iniciar Búsqueda RSS", font=("Century Gothic", 14, "bold"), 
                                     fg_color="#333333", hover_color="#555555", command=self.start_thread, height=40)
        self.btn_run.grid(row=1, column=0, padx=15, pady=15, sticky="w")
        
        self.status_label = ctk.CTkLabel(self.controls_frame, text="Estado: Esperando inicio...", font=("Century Gothic", 13), text_color="#aaaaaa")
        self.status_label.grid(row=1, column=1, columnspan=4, sticky="e", padx=15)

        self.progress_bar = ctk.CTkProgressBar(self.controls_frame, mode="determinate", progress_color="#00fa9a")
        self.progress_bar.grid(row=2, column=0, columnspan=5, padx=15, pady=(0, 15), sticky="ew")
        self.progress_bar.set(0)

        # Consola Log
        self.log_box = ctk.CTkTextbox(self, font=("Consolas", 12), fg_color="#0d0d0d", text_color="#00fa9a")
        self.log_box.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="nsew")
        self.log_box.configure(state="disabled")

        self.original_stdout = sys.stdout
        self.refresh_analysis_status()

    def get_selected_sources(self):
        selected_sources = []
        if self.el_tiempo_var.get():
            selected_sources.append("El Tiempo")
        if self.portafolio_var.get():
            selected_sources.append("Portafolio")
        return selected_sources

    def on_source_toggle(self):
        if self.get_selected_sources():
            self.btn_run.configure(state="normal")
            self.status_label.configure(text="Estado: Esperando inicio...", text_color="#aaaaaa")
        else:
            self.btn_run.configure(state="disabled")
            self.status_label.configure(text="Estado: Seleccione al menos un medio", text_color="#ffb347")

    def refresh_analysis_status(self):
        if get_groq_api_key():
            self.analysis_label.configure(
                text="Análisis El Tiempo: Groq AI activo (3 palabras)",
                text_color="#00fa9a"
            )
            return "Groq AI"

        self.analysis_label.configure(
            text="Análisis El Tiempo: modo heurístico",
            text_color="#ffb347"
        )
        return "Heuristico"

    def update_progress(self, pct):
        self.progress_bar.set(pct / 100.0)
        if pct == 100:
            self.status_label.configure(text="Estado: Proceso completado exitosamente", text_color="#00fa9a")
            self.btn_run.configure(state="normal", text="Búsqueda Finalizada")

    def start_thread(self):
        selected_sources = self.get_selected_sources()
        if not selected_sources:
            self.status_label.configure(text="Estado: Seleccione al menos un medio", text_color="#ffb347")
            return
        include_amp = self.amp_var.get()
        analysis_mode = self.refresh_analysis_status()
        self.btn_run.configure(state="disabled", text="Procesando...")
        amp_txt = "con AMP" if include_amp else "sin AMP"
        self.status_label.configure(text=f"Estado: Extrayendo ({analysis_mode}, {amp_txt})...", text_color="#FFFFFF")
        self.progress_bar.set(0.0)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        
        # Redireccionar print() -> UI
        sys.stdout = RealtimeLog(self.log_box, self.update_progress) # type: ignore
        
        # Iniciar hilo en segundo plano
        self.thread = threading.Thread(target=self.run_process, args=(selected_sources, include_amp), daemon=True)
        self.thread.start()

    def run_process(self, selected_sources, include_amp=True):
        try:
            run_scraper_selected(selected_sources, include_amp=include_amp)
        except Exception as e:
            print(f"\n[ERROR] Ocurrió un fallo: {e}")
            self.status_label.configure(text="Estado: Error en la ejecución", text_color="#FF0000")
            self.btn_run.configure(state="normal", text="Reintentar")
        finally:
            sys.stdout = self.original_stdout
            
    def on_closing(self):
        try:
            sys.stdout = self.original_stdout
        except:
            pass
        self.destroy()

if __name__ == "__main__":
    if HAS_GUI and ctk is not None:
        try:
            app = FlowcatsApp()
            app.protocol("WM_DELETE_WINDOW", app.on_closing)
            app.mainloop()
        except Exception as e:
            print(f"[!] Error al iniciar interfaz gráfica: {e}")
            print("[*] Ejecutando extracción en modo consola (CLI)...")
            run_scraper_selected(["El Tiempo", "Portafolio"], include_amp=True)
    else:
        print("[!] Interfaz gráfica no disponible en este entorno (falta _tkinter / CustomTkinter).")
        print("[*] Ejecutando extracción en modo consola (CLI)...")
        run_scraper_selected(["El Tiempo", "Portafolio"], include_amp=True)
