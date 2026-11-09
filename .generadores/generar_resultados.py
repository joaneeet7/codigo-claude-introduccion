"""Genera '2-Resultados/' de las clases que reutilizan demos del curso Claude
Associate Foundations, con la cabecera y la numeración de este curso.

Uso: python3 -I .generadores/generar_resultados.py   (macOS, usa Google Chrome headless)
- Compara modelos con la misma tarea ← Selección de productos y modelos, pasos 3 y 4 (resultados/datos).
- Monta un proyecto para tu día a día ← Configuración y gestión del conocimiento, pasos 1 y 2 (resultados/html).
- Tu checklist de uso responsable ← Gobernanza, riesgo y uso responsable, paso 2 (resultados/html).
Los resultados de las demás prácticas se añadirán cuando se graben.
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "resultados"))
from comun import CHECKLIST, CHROME, MODELOS, PROYECTO, carpeta  # noqa: E402
import pdf_resultado  # noqa: E402

DEST = HERE.parent / "2-Resultados"
JSON = {MODELOS: ["Paso 1 - Tarea simple, modelo ligero frente a potente", "Paso 2 - Tarea compleja, abrir en A o en B"]}
HTML = {
    (PROYECTO, "proyecto_paso1.html"): "Paso 1 - Instrucciones vagas.pdf",
    (PROYECTO, "proyecto_paso2.html"): "Paso 2 - Instrucciones efectivas y conocimiento.pdf",
    (CHECKLIST, "checklist_paso3.html"): "Paso 3 - Decisión sobre personas.pdf",
}


def main():
    for practica, nombres in JSON.items():
        for nombre in nombres:
            pdf_resultado.generar(str(HERE / "resultados" / "datos" / f"{nombre}.json"), DEST / carpeta(practica))
    for (practica, src), nombre in HTML.items():
        out = DEST / carpeta(practica) / nombre
        out.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run([CHROME, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={out}",
                        (HERE / "resultados" / "html" / src).as_uri()], check=True, capture_output=True)
        print(nombre)


if __name__ == "__main__":
    main()
