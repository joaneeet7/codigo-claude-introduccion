"""Utilidades compartidas por los generadores del curso Introducción a Claude."""
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MATERIALES = ROOT / "0-Materiales"
ENUNCIADOS = ROOT / "1-Enunciados"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# Prácticas del curso, agrupadas por sección del temario. Sin números de clase:
# la numeración del programa puede cambiar. Cada práctica se identifica por su
# nombre, que se usa en las carpetas de materiales y resultados y en el PDF del
# enunciado: <Sección>/<Práctica>.
PETICION = "De una petición vaga a un resultado profesional"
MODELOS = "Compara modelos con la misma tarea"
INFORME_PDF = "Informe ejecutivo a partir de un PDF"
GASTOS = "Analiza una hoja de gastos"
GUIA = "Crea una guía visual para tu equipo"
MERCADO = "Informe de mercado con fuentes verificadas"
PROPUESTA = "Redacta una propuesta profesional con Claude"
CONVERSACIONES = "Ensaya conversaciones difíciles"
PROYECTO = "Monta un proyecto para tu día a día"
SEMANA = "Prepara tu semana con tu correo y tu calendario"
WEBS = "Compara opciones en varias webs"
SKILL = "Crea tu primera skill"
COWORK = "Delega una tarea completa en Cowork"
APP = "Construye tu primera aplicación con Claude Code"
CHECKLIST = "Tu checklist de uso responsable"

SECCIONES = {
    "Primeros pasos con Claude": [PETICION, MODELOS],
    "Documentos, datos y contenido visual": [INFORME_PDF, GASTOS, GUIA],
    "Investigación y escritura": [MERCADO, PROPUESTA, CONVERSACIONES],
    "Organización y conectores": [PROYECTO, SEMANA],
    "Herramientas y automatización": [WEBS, SKILL, COWORK, APP],
    "Privacidad y uso responsable": [CHECKLIST],
}
SECCION = {p: s for s, ps in SECCIONES.items() for p in ps}


def carpeta(practica):
    """Ruta relativa <Sección>/<Práctica>."""
    return Path(SECCION[practica]) / practica


def html_a_pdf(html, destino, esperar=False):
    """Imprime un HTML a PDF con Chrome headless."""
    tmp = Path(tempfile.mkdtemp(prefix="intro_claude_"))
    try:
        src, pdf = tmp / "d.html", tmp / "d.pdf"
        src.write_text(html, encoding="utf-8")
        cmd = [CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={pdf}"]
        if esperar:
            cmd.insert(3, "--virtual-time-budget=3000")
        subprocess.run(cmd + [src.as_uri()], check=True, capture_output=True)
        destino.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(pdf, destino)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
