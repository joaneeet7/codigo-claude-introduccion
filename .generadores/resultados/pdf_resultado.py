"""Genera un PDF 16:9 (1920 × 1080) de resultados de demo a partir de un .json.

Copiado del curso Claude Associate Foundations para reutilizar sus resultados.
Lo usa generar_resultados.py; no hace falta llamarlo directamente.

Izquierda: lo que se hizo y los resultados extraídos (métricas y extractos).
Derecha: conclusiones con tabla comparativa, puntos numerados y frase clave.
En las celdas de la tabla, "✓", "✕" y "~" al inicio se pintan en verde, rojo
y ámbar hasta la primera coma o dos puntos. En conclusiones y frase clave,
**texto** va en negrita y *texto* en el color de acento.
"""
import html
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
DATOS = HERE / "datos"
REPO = HERE.parent.parent
DEST = REPO / "2-Resultados"
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

CSS = """
@page { size: 297mm 167.06mm; margin: 0; }  /* 16:9, como las grabaciones 1920 × 1080 */
:root { --ink: #1f1d1a; --muted: #6b665f; --hair: #e3ddd3; --accent: #c15f3c;
  --paper: #fbf8f3; --ok: #2f8f5b; --ko: #c8463d; --mid: #c98a1b; }
* { box-sizing: border-box; font-variant-ligatures: none; }
html, body { margin: 0; background: #fff; color: var(--ink);
  font-family: 'Avenir Next', 'Helvetica Neue', sans-serif; }
.page { width: 297mm; height: 167.06mm; padding: 10mm 15mm 8mm; display: flex;
  flex-direction: column; break-after: page; overflow: hidden; }
.head { display: flex; justify-content: space-between; align-items: flex-end;
  border-bottom: 1.5px solid var(--ink); padding-bottom: 4mm; }
.label { font-size: 8.5pt; letter-spacing: .22em; text-transform: uppercase;
  color: var(--accent); font-weight: 600; display: flex; align-items: center; gap: 3mm; }
.label::before { content: ""; width: 9mm; height: 1.5px; background: var(--accent); }
h1 { font-size: 22pt; line-height: 1.1; margin: 2.5mm 0 0; font-weight: 600; letter-spacing: -.01em; }
.head .meta { font-size: 9pt; color: var(--muted); text-align: right; line-height: 1.5; }
.cols { flex: 1; display: grid; grid-template-columns: 42fr 58fr; gap: 11mm; padding-top: 5mm; min-height: 0; }
.col { min-height: 0; overflow: hidden; }
.inner { display: flex; flex-direction: column; transform-origin: top left; }
.left { background: var(--paper); margin: -5mm 0 -8mm -15mm; padding: 5mm 9mm 0 15mm; }
h2 { margin: 0 0 3.5mm; }
.block { margin-bottom: 5mm; }
.ctx { font-size: 10.5pt; line-height: 1.45; color: var(--ink); }
.prompt { font-family: 'SF Mono', Menlo, monospace; font-size: 8.6pt; line-height: 1.45;
  border-left: 2px solid var(--accent); padding: 1mm 0 1mm 3.5mm; color: #3a3631;
  margin-top: 2.5mm; white-space: pre-wrap; }
.metrics { display: flex; gap: 5mm; border-top: 1px solid var(--hair);
  border-bottom: 1px solid var(--hair); padding: 3.5mm 0; }
.metric { flex: 1; }
.metric b { display: block; font-size: 21pt; line-height: 1.05; font-weight: 600; }
.metric span { font-size: 7.8pt; letter-spacing: .1em; text-transform: uppercase; color: var(--muted); }
.res { list-style: none; margin: 0; padding: 0; }
.res li { font-size: 10pt; line-height: 1.42; padding: 2mm 0; border-bottom: 1px solid var(--hair); }
.res li b { display: block; font-size: 8pt; letter-spacing: .12em; text-transform: uppercase;
  color: var(--muted); font-weight: 600; margin-bottom: .6mm; }
.quote { font-size: 10pt; line-height: 1.45; font-style: italic; color: #3a3631; margin: 0 0 2.5mm; }
.quote::before { content: "“"; color: var(--accent); font-weight: 700; }
.quote::after { content: "”"; color: var(--accent); font-weight: 700; }
table { width: 100%; border-collapse: collapse; font-size: 9.6pt; }
th { text-align: left; font-size: 7.6pt; letter-spacing: .12em; text-transform: uppercase;
  color: var(--muted); font-weight: 600; padding: 0 2.5mm 2mm 0; border-bottom: 1.5px solid var(--ink); }
td { padding: 2.2mm 2.5mm 2.2mm 0; border-bottom: 1px solid var(--hair); vertical-align: top; line-height: 1.35; }
td:first-child { font-weight: 600; }
.ok { color: var(--ok); font-weight: 600; } .ko { color: var(--ko); font-weight: 600; }
.mid { color: var(--mid); font-weight: 600; }
.concl { list-style: none; margin: 0; padding: 0; }
.concl li { display: grid; grid-template-columns: 10mm 1fr; font-size: 10.4pt; line-height: 1.45;
  padding: 2.4mm 0; border-bottom: 1px solid var(--hair); }
.concl li span { color: var(--accent); font-weight: 700; font-size: 9.5pt; }
.key { font-size: 14pt; font-weight: 600; line-height: 1.3; margin-top: auto; padding: 4mm 0 2mm; }
.key em, .concl em { font-style: normal; color: var(--accent); }
.right .inner { min-height: 100%; }
"""

# Encoge cada columna hasta que quepa y deja la escala mínima en <body data-scale>.
FIT = """
(function () {
  let zmin = 1;
  document.querySelectorAll('.col').forEach(c => {
    const inner = c.firstElementChild;
    let z = 1;
    while (inner.scrollHeight * z > c.clientHeight + 0.5 && z > 0.5) {
      z -= 0.01;
      inner.style.transform = 'scale(' + z + ')';
      inner.style.width = (100 / z) + '%';
    }
    zmin = Math.min(zmin, z);
  });
  document.body.dataset.scale = zmin.toFixed(2);
  document.querySelectorAll('script').forEach(s => s.remove());
})();
"""


def esc(s):
    return html.escape(str(s))


def marcado(s):
    """**negrita** y *acento* sobre texto ya escapado."""
    s = esc(s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", s)
    return re.sub(r"\*(.+?)\*", r"<em>\1</em>", s)


def celda(s):
    """Pinta ✓ / ✕ / ~ iniciales y el texto hasta la primera coma o dos puntos."""
    s = str(s)
    cls = {"✓": "ok", "✕": "ko", "~": "mid"}.get(s[:1])
    if not cls:
        return esc(s)
    m = re.match(r"([^,:]*)(.*)", s, re.S)
    return f'<span class="{cls}">{esc(m.group(1))}</span>{esc(m.group(2))}'


def label(t):
    return f'<h2 class="label">{esc(t)}</h2>'


def construir(d):
    meta = "<br>".join(esc(m) for m in d.get("meta", []))
    etiqueta = re.sub(r"^\s*Secci[oó]n\s*\d+\s*·\s*", "", d.get("etiqueta", "Resultados de la demo"))
    left, right = [], []
    if d.get("contexto") or d.get("prompt"):
        b = label(d.get("etiqueta_contexto", "Qué le pedimos a Claude"))
        if d.get("contexto"):
            b += f'<div class="ctx">{esc(d["contexto"])}</div>'
        if d.get("prompt"):
            b += f'<div class="prompt">{html.escape(d["prompt"])}</div>'
        left.append(f'<div class="block">{b}</div>')
    if d.get("metricas"):
        m = "".join(f'<div class="metric"><b>{esc(x["valor"])}</b><span>{esc(x["etiqueta"])}</span></div>'
                    for x in d["metricas"])
        left.append(f'<div class="block">{label(d.get("etiqueta_metricas", "Resultados"))}<div class="metrics">{m}</div></div>')
    if d.get("resultados"):
        r = "".join(f'<li><b>{esc(x["etiqueta"])}</b>{esc(x["texto"])}</li>' for x in d["resultados"])
        left.append(f'<div class="block">{label(d.get("etiqueta_resultados", "Lo que ha devuelto"))}<ul class="res">{r}</ul></div>')
    if d.get("extractos"):
        q = "".join(f'<p class="quote">{esc(x)}</p>' for x in d["extractos"])
        left.append(f'<div class="block">{label(d.get("etiqueta_extractos", "Extractos de la respuesta"))}{q}</div>')
    if d.get("tabla"):
        t = d["tabla"]
        th = "".join(f"<th>{esc(c)}</th>" for c in t["columnas"])
        tr = "".join("<tr>" + "".join(f"<td>{celda(c)}</td>" for c in f) + "</tr>" for f in t["filas"])
        right.append(f'<div class="block">{label(t.get("titulo", "Qué ha cambiado"))}'
                     f'<table><thead><tr>{th}</tr></thead><tbody>{tr}</tbody></table></div>')
    if d.get("conclusiones"):
        c = "".join(f'<li><span>{i:02d}</span><div>{marcado(x)}</div></li>'
                    for i, x in enumerate(d["conclusiones"], 1))
        right.append(f'<div class="block">{label("Conclusiones")}<ol class="concl">{c}</ol></div>')
    if d.get("frase_clave"):
        right.append(f'<div class="key">{marcado(d["frase_clave"])}</div>')
    return f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>{esc(d["titulo"])}</title><style>{CSS}</style></head><body><div class="page">
<div class="head"><div><div class="label">{esc(etiqueta)}</div>
<h1>{esc(d["titulo"])}</h1></div><div class="meta">{meta}</div></div>
<div class="cols"><div class="col left"><div class="inner">{"".join(left)}</div></div>
<div class="col right"><div class="inner">{"".join(right)}</div></div></div>
</div><script>{FIT}</script></body></html>"""


def chrome(*args):
    return subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                           "--virtual-time-budget=4000", *args],
                          capture_output=True, text=True, check=True).stdout


def generar(json_path, out_dir=None):
    json_path = Path(json_path)
    if not json_path.is_absolute():
        json_path = DATOS / json_path
    d = json.loads(json_path.read_text(encoding="utf-8"))
    out_dir = Path(out_dir) if out_dir else DEST / json_path.parent.name
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / (json_path.stem + ".pdf")
    tmp = Path(tempfile.mkdtemp(prefix="resultado_"))
    try:
        src = tmp / "in.html"
        src.write_text(construir(d), encoding="utf-8")
        # 1) Chrome ejecuta el ajuste y devuelve el DOM ya encogido; 2) se imprime ese DOM.
        dom = chrome("--dump-dom", src.as_uri())
        escala = float(re.search(r'data-scale="([\d.]+)"', dom).group(1))
        fijo = tmp / "fijo.html"
        fijo.write_text("<!doctype html>" + dom, encoding="utf-8")
        chrome(f"--print-to-pdf={out}", fijo.as_uri())
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"{out.name}: escala {escala:g}")
    if escala < 0.85:
        print("AVISO: texto reducido por debajo del 85 %; conviene recortar contenido.")
    return out


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    salida = None
    args = sys.argv[1:]
    if "--salida" in args:
        i = args.index("--salida")
        salida = args[i + 1]
        del args[i:i + 2]
    for a in args:
        generar(a, salida)
