"""Utilidades compartidas por los generadores del curso Introducción a Claude."""
import shutil
import subprocess
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
MATERIALES = ROOT / "0-Materiales"
ENUNCIADOS = ROOT / "1-Enunciados"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

# Nombre de cada clase práctica: se usa en las carpetas de materiales,
# resultados y en el nombre del PDF del enunciado.
CLASES = {
    7: "De una petición vaga a un resultado profesional",
    9: "Compara modelos con la misma tarea",
    12: "Informe ejecutivo a partir de un PDF",
    14: "Analiza una hoja de gastos",
    16: "Crea una guía visual para tu equipo",
    18: "Informe de mercado con fuentes verificadas",
    20: "Redacta una propuesta profesional con Claude",
    22: "Ensaya conversaciones difíciles",
    24: "Monta un proyecto para tu día a día",
    26: "Prepara tu semana con tu correo y tu calendario",
    28: "Compara opciones en varias webs",
    30: "Crea tu primera skill",
    32: "Delega una tarea completa en Cowork",
    34: "Construye tu primera aplicación con Claude Code",
    37: "Tu checklist de uso responsable",
}


def carpeta(n):
    return f"Clase {n:02d} - {CLASES[n]}"


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
