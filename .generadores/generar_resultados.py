"""Genera '2-Resultados/' de las clases que reutilizan demos del curso Claude
Associate Foundations, con la cabecera y la numeración de este curso.

Uso: python3 -I .generadores/generar_resultados.py   (macOS, usa Google Chrome headless)
- Clase 09 ← Selección de productos y modelos, pasos 3 y 4 (datos en resultados/datos).
- Clase 24 ← Configuración y gestión del conocimiento, pasos 1 y 2 (resultados/html).
- Clase 37 ← Gobernanza, riesgo y uso responsable, paso 2 (resultados/html).
Los resultados de las demás clases se añadirán cuando se graben.
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "resultados"))
from comun import CHROME, carpeta  # noqa: E402
import pdf_resultado  # noqa: E402

DEST = HERE.parent / "2-Resultados"
JSON = {9: ["Paso 1 - Tarea simple, modelo ligero frente a potente", "Paso 2 - Tarea compleja, abrir en A o en B"]}
HTML = {
    (24, "clase24_paso1.html"): "Paso 1 - Instrucciones vagas.pdf",
    (24, "clase24_paso2.html"): "Paso 2 - Instrucciones efectivas y conocimiento.pdf",
    (37, "clase37_paso3.html"): "Paso 3 - Decisión sobre personas.pdf",
}


def main():
    for n, nombres in JSON.items():
        for nombre in nombres:
            pdf_resultado.generar(str(HERE / "resultados" / "datos" / f"{nombre}.json"), DEST / carpeta(n))
    for (n, src), nombre in HTML.items():
        out = DEST / carpeta(n) / nombre
        out.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run([CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={out}",
                        (HERE / "resultados" / "html" / src).as_uri()], check=True, capture_output=True)
        print(nombre)


if __name__ == "__main__":
    main()
