"""Genera '0-Materiales/' del curso Introducción a Claude.

Uso: python3 -I .generadores/generar_materiales.py   (macOS, usa Google Chrome headless)
Todos los datos son ficticios. Los materiales de «Compara modelos con la misma
tarea» y «Monta un proyecto para tu día a día», y los CV de «Tu checklist de uso
responsable», se reutilizan del curso Claude Associate Foundations: están copiados tal
cual en 0-Materiales y este script no los toca. El script es determinista: las cifras de control
que imprime al final son las que aparecen en los enunciados (no las cambies sin
actualizar también generar_enunciados.py).
"""
import base64
import datetime as dt
import io
import random
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from comun import COWORK, GASTOS, INFORME_PDF, MATERIALES, MERCADO, SEMANA, SKILL, carpeta, html_a_pdf  # noqa: E402
from materiales_texto import TEXTOS  # noqa: E402

import docx  # noqa: E402
from docx.shared import Pt  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

CONTROL = {}


def ruta(practica, nombre):
    p = MATERIALES / carpeta(practica) / nombre
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def eur(x):
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return s + " €"


def hoja(ws, cabecera, filas, anchos):
    ws.append(cabecera)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E5F")
        c.alignment = Alignment(vertical="center")
    for f in filas:
        ws.append(f)
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


# ------------------------------------------------------------------ textos
def textos():
    for (practica, nombre), contenido in TEXTOS.items():
        ruta(practica, nombre).write_text(contenido, encoding="utf-8")


# ------------------------------------------------------------------ Informe ejecutivo a partir de un PDF
CSS_INFORME = """
@page { size: A4; margin: 20mm 20mm 18mm; }
body { font-family: 'Helvetica Neue', Arial, sans-serif; font-size: 10.5pt; line-height: 1.5; color: #222; }
h1 { font-size: 26pt; margin: 0 0 4mm; color: #6b3e1f; }
h2 { font-size: 14pt; color: #6b3e1f; border-bottom: 1px solid #d8c7b5; padding-bottom: 2mm; margin-top: 0; }
.portada { height: 250mm; display: flex; flex-direction: column; justify-content: center; break-after: page; }
.portada .sub { font-size: 14pt; color: #555; }
.portada .pie { margin-top: 30mm; font-size: 10pt; color: #777; }
section { break-after: page; }
table { border-collapse: collapse; width: 100%; margin: 4mm 0; font-size: 10pt; }
th, td { border-bottom: 1px solid #ddd; padding: 2mm 3mm; text-align: right; }
th:first-child, td:first-child { text-align: left; }
th { background: #f3ece4; }
img { width: 100%; margin: 3mm 0; }
.nota { font-size: 9pt; color: #666; }
"""


def grafico_mermas():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    v25 = [6.1, 5.8, 6.0, 6.2, 5.9, 6.3, 6.5, 6.4, 6.0, 6.1, 5.9, 6.2]
    v26 = [5.9, 6.0, 5.7, 6.1, 6.0, 6.2, 6.4, 6.6, 9.8, 10.9, 11.4, 10.7]
    fig, ax = plt.subplots(figsize=(8, 3.6), dpi=160)
    x = range(12)
    ax.bar([i - 0.2 for i in x], v25, width=0.4, label="2025", color="#c9b39b")
    ax.bar([i + 0.2 for i in x], v26, width=0.4, label="2026", color="#6b3e1f")
    ax.set_xticks(list(x), meses)
    ax.set_ylabel("% de pan no vendido")
    ax.set_title("Mermas mensuales: pan y bollería no vendidos (% de la producción)")
    ax.set_ylim(0, 13)
    ax.legend(frameon=False)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def informe_arrabal():
    tiendas = [("Centro", 1_052_000, 968_000), ("Ruzafa", 0, 512_000), ("Benimaclet", 615_000, 652_000),
               ("Campanar", 589_000, 618_000), ("Patraix", 561_000, 590_000), ("Cabanyal", 571_000, 609_000)]
    online = (420_000, 481_000)
    b2b = (387_000, 0)
    filas = "".join(f"<tr><td>{t}</td><td>{eur(a)[:-5] + ' €' if a else '—'}</td><td>{eur(b)[:-5]} €</td>"
                    f"<td>{(f'{(b - a) / a * 100:+.1f} %'.replace('.', ',')) if a else 'apertura'}</td></tr>"
                    for t, a, b in tiendas)
    tot25 = sum(a for _, a, _ in tiendas) + online[0] + b2b[0]
    b2b26 = 4_820_000 - sum(b for _, _, b in tiendas) - online[1]
    tot26 = sum(b for _, _, b in tiendas) + online[1] + b2b26
    assert tot26 == 4_820_000
    CONTROL["Informe PDF · ventas 2026"] = f"{eur(tot26)} ({(tot26 - tot25) / tot25 * 100:+.1f} % frente a {eur(tot25)})"
    img = grafico_mermas()
    html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Informe anual 2026 · Panes Arrabal</title>
<style>{CSS_INFORME}</style></head><body>
<div class="portada"><h1>Informe anual 2026</h1><div class="sub">Panes Arrabal · Obrador y tiendas de pan artesano</div>
<div class="pie">Documento interno para el comité de dirección · Valencia, enero de 2027<br>Datos ficticios para uso formativo.</div></div>

<section><h2>1. Carta de la dirección</h2>
<p>2026 ha sido el año de mayor crecimiento de Panes Arrabal desde que abrimos el obrador. Hemos vendido 4,82 millones de euros, con la apertura de nuestra sexta tienda en Ruzafa en el mes de abril, que ha superado las previsiones desde el primer trimestre.</p>
<p>El esfuerzo del equipo ha sido enorme. En septiembre cambiamos el horario de la hornada principal, que pasó de las 4:00 a las 6:00 de la mañana, para mejorar las condiciones de trabajo del obrador y reducir el gasto en horas nocturnas. La medida ha sido muy bien recibida por la plantilla.</p>
<p>Afrontamos 2027 con prudencia: el precio de la harina sigue subiendo y tendremos que renegociar el alquiler de la tienda Centro. Pero contamos con un equipo comprometido y una clientela fiel.</p>
<p><b>Carmen Lloret</b>, directora general</p></section>

<section><h2>2. Ventas por tienda</h2>
<table><tr><th>Tienda</th><th>2025</th><th>2026</th><th>Variación</th></tr>{filas}</table>
<p>Cinco de las seis tiendas crecen. La tienda Centro baja un 8 %, sobre todo a partir de junio, por las obras de peatonalización de su calle, que han reducido el paso de gente. El ayuntamiento prevé terminarlas en febrero de 2027.</p>
<p>Ruzafa, abierta en abril, ha vendido 512.000 euros en solo nueve meses, por encima de lo previsto en el plan de apertura.</p></section>

<section><h2>3. Canales de venta</h2>
<table><tr><th>Canal</th><th>2025</th><th>2026</th></tr>
<tr><td>Tiendas</td><td>{eur(sum(a for _, a, _ in tiendas))}</td><td>{eur(sum(b for _, _, b in tiendas))}</td></tr>
<tr><td>Pedidos online con recogida</td><td>{eur(online[0])}</td><td>{eur(online[1])}</td></tr>
<tr><td>Cafeterías y restaurantes (B2B)</td><td>{eur(b2b[0])}</td><td>{eur(b2b26)}</td></tr>
<tr><td><b>Total</b></td><td><b>{eur(tot25)}</b></td><td><b>{eur(tot26)}</b></td></tr></table>
<p>Los pedidos online con recogida en tienda crecen un 14,5 % y ya suponen el 10 % de las ventas. El canal B2B se mantiene estable con 23 clientes activos, aunque perdimos en octubre la cuenta de la cadena Café Volta, que representaba unos 40.000 euros al año.</p></section>

<section><h2>4. Producción y mermas</h2>
<p>El obrador ha producido una media de 9.400 piezas diarias. Hemos incorporado tres referencias nuevas de pan de masa madre y renovado la cámara de fermentación.</p>
<p>Las mermas se han mantenido bajo control durante todo el año, en línea con el objetivo de no superar el 7 % de la producción.</p>
<img src="data:image/png;base64,{img}" alt="Gráfico de mermas mensuales 2025 y 2026">
<p class="nota">Fuente: registro de producción y devoluciones de tiendas.</p></section>

<section><h2>5. Personas</h2>
<p>Cerramos el año con 64 personas en plantilla (58 en 2025). La rotación ha sido del 18 %, concentrada en las tiendas, y hemos impartido 1.150 horas de formación, sobre todo en atención al cliente y seguridad alimentaria.</p>
<p>La encuesta de clima interno sube de 6,9 a 7,6 sobre 10. La mejor valorada es la nueva hora de la hornada.</p></section>

<section style="break-after: auto"><h2>6. Riesgos y prioridades para 2027</h2>
<ul><li><b>Harina:</b> el proveedor principal ha anunciado una subida del 14 % a partir de febrero.</li>
<li><b>Alquiler de la tienda Centro:</b> el contrato vence el 31 de marzo de 2027 y la propiedad ha pedido una subida del 20 %.</li>
<li><b>Aplicación de pedidos:</b> queremos lanzar en el segundo trimestre una aplicación propia para los pedidos con recogida.</li>
<li><b>Canal B2B:</b> recuperar volumen tras la pérdida de Café Volta, con el objetivo de llegar a 28 clientes activos.</li></ul>
</section></body></html>"""
    html_a_pdf(html, ruta(INFORME_PDF, "informe_anual_2026_panes_arrabal.pdf"))


# ------------------------------------------------------------------ Analiza una hoja de gastos
def gastos_lumbre():
    rnd = random.Random(14)
    filas = []

    def f(d, concepto, prov, cat, imp, pago):
        filas.append([d, concepto, prov, cat, round(imp, 2), pago])

    for m in (7, 8, 9):
        f(dt.date(2026, m, 1), "Alquiler oficina", "Inmobiliaria Portal Norte", "Alquiler y suministros", 1850.00, "Transferencia")
        f(dt.date(2026, m, 3), "Electricidad", "Energía Clara", "Alquiler y suministros", {7: 212.40, 8: 268.15, 9: 231.90}[m], "Domiciliación")
        f(dt.date(2026, m, 3), "Internet y teléfono", "Telecom Viva", "Alquiler y suministros", 89.90, "Domiciliación")
        f(dt.date(2026, m, 5), "Adobe Creative Cloud (3 licencias)", "Adobe", "Software y suscripciones", 251.67, "Tarjeta empresa")
        f(dt.date(2026, m, 6), "Licencia Archicad", "Graphisoft", "Software y suscripciones", 312.00, "Tarjeta empresa")
        f(dt.date(2026, m, 8), "Google Workspace (7 usuarios)", "Google", "Software y suscripciones", 64.80, "Tarjeta empresa")
        f(dt.date(2026, m, 28), "Gestoría y nóminas", "Asesoría Ferrer", "Servicios profesionales", 290.40, "Transferencia")
        f(dt.date(2026, m, 30), "Limpieza oficina", "Limpiezas Brillo", "Servicios profesionales", 180.00, "Transferencia")
        f(dt.date(2026, m, 15), "Campaña LinkedIn", "LinkedIn", "Marketing", {7: 300.00, 8: 150.00, 9: 300.00}[m], "Tarjeta empresa")
    # Nuevas suscripciones: el software crece
    f(dt.date(2026, 8, 12), "Render en la nube", "RenderNube", "Software y suscripciones", 149.00, "Tarjeta empresa")
    f(dt.date(2026, 9, 12), "Render en la nube", "RenderNube", "Software y suscripciones", 149.00, "Tarjeta empresa")
    f(dt.date(2026, 9, 18), "Gestor de proyectos (7 usuarios)", "Planora", "Software y suscripciones", 96.00, "Tarjeta empresa")
    # Material de oficina y maquetas
    for m, n in ((7, 6), (8, 3), (9, 6)):
        for _ in range(n):
            d = dt.date(2026, m, rnd.randint(2, 27))
            concepto, prov, lo, hi = rnd.choice([
                ("Papel y cartulinas", "Papelería Trazo", 18, 65), ("Tóner impresora", "OfiMax", 79, 139),
                ("Material de maquetas", "Maquetas y Más", 35, 180), ("Cuchillas y cola", "Bellas Artes Goya", 12, 40)])
            f(d, concepto, prov, "Material de oficina", rnd.uniform(lo, hi), "Tarjeta empresa")
    f(dt.date(2026, 7, 9), "Silla ergonómica", "OfiMax", "Material de oficina", 189.00, "Tarjeta empresa")
    f(dt.date(2026, 8, 4), "Devolución silla ergonómica", "OfiMax", "Material de oficina", -189.00, "Tarjeta empresa")
    # Desplazamientos (algunos mal etiquetados como «Viajes»)
    for m, n in ((7, 5), (8, 3), (9, 5)):
        for _ in range(n):
            d = dt.date(2026, m, rnd.randint(2, 28))
            concepto, prov, lo, hi = rnd.choice([
                ("Taxi visita a obra", "Taxi Madrid", 14, 32), ("Parking visita a obra", "Parking Plaza", 6, 18),
                ("Gasolina coche estudio", "Estación Repsol Castellana", 45, 72), ("Cercanías visita a obra", "Renfe", 3.4, 8.6)])
            f(d, concepto, prov, "Desplazamientos", rnd.uniform(lo, hi), "Tarjeta empresa")
    f(dt.date(2026, 7, 14), "Tren a Toledo (visita a obra)", "Renfe", "Viajes", 42.60, "Tarjeta empresa")
    f(dt.date(2026, 7, 22), "Tren a Segovia (reunión cliente)", "Renfe", "Viajes", 38.20, "Tarjeta empresa")
    f(dt.date(2026, 8, 19), "Tren a Valladolid (visita a obra)", "Renfe", "Viajes", 56.80, "Tarjeta empresa")
    # Feria en Barcelona: el pico de septiembre
    f(dt.date(2026, 9, 21), "AVE Madrid-Barcelona, ida y vuelta (3 personas), feria Construmat", "Renfe", "Desplazamientos", 642.30, "Tarjeta empresa")
    f(dt.date(2026, 9, 22), "Hotel Barcelona, 2 noches (3 habitaciones), feria Construmat", "Hotel Diagonal Mar", "Desplazamientos", 1026.00, "Tarjeta empresa")
    f(dt.date(2026, 9, 23), "Taxis Barcelona, feria Construmat", "Taxi Barcelona", "Desplazamientos", 84.60, "Tarjeta empresa")
    # Comidas con clientes
    for m, n in ((7, 2), (8, 1), (9, 3)):
        for _ in range(n):
            d = dt.date(2026, m, rnd.randint(2, 28))
            prov = rnd.choice(["Restaurante La Tasquita", "Casa Pepa", "Bistró Olivar", "Taberna del Arco"])
            f(d, "Comida con cliente", prov, "Comidas con clientes", rnd.uniform(48, 135), "Tarjeta empresa")
    f(dt.date(2026, 8, 25), "Curso BIM avanzado (2 personas)", "Escuela Técnica Forma", "Formación", 450.00, "Transferencia")
    f(dt.date(2026, 7, 17), "Impresión dossier de proyectos", "Imprenta Gutenberg", "Marketing", 210.00, "Transferencia")
    filas.sort(key=lambda r: (r[0], r[1]))
    # Cargo duplicado: la misma suscripción registrada dos veces el mismo día
    dup = next(i for i, r in enumerate(filas) if r[0] == dt.date(2026, 9, 5) and r[2] == "Adobe")
    filas.insert(dup + 1, list(filas[dup]))
    total_hoja = round(sum(r[4] for r in filas), 2)
    total_banco = round(total_hoja - filas[dup][4], 2)
    CONTROL["Gastos · filas de la hoja"] = len(filas)
    CONTROL["Gastos · total de la hoja"] = eur(total_hoja)
    CONTROL["Gastos · total del extracto (sin el duplicado)"] = eur(total_banco)
    por_mes = {m: round(sum(r[4] for r in filas if r[0].month == m and r[3] in ("Desplazamientos", "Viajes")), 2) for m in (7, 8, 9)}
    CONTROL["Gastos · desplazamientos + viajes por mes"] = {m: eur(v) for m, v in por_mes.items()}
    sw = {m: round(sum(r[4] for r in filas if r[0].month == m and r[3] == "Software y suscripciones"), 2) for m in (7, 8, 9)}
    CONTROL["Gastos · software por mes (con duplicado)"] = {m: eur(v) for m, v in sw.items()}

    wb = Workbook()
    ws = wb.active
    ws.title = "Gastos T3 2026"
    hoja(ws, ["Fecha", "Concepto", "Proveedor", "Categoría", "Importe (€)", "Pagado con"], filas, [12, 58, 28, 26, 13, 17])
    for r in ws.iter_rows(min_row=2):
        r[0].number_format = "DD/MM/YYYY"
        r[4].number_format = '#,##0.00'
    wb.save(ruta(GASTOS, "gastos_T3_2026_estudio_lumbre.xlsx"))
    return CONTROL["Gastos · total del extracto (sin el duplicado)"]


def ticket_restaurante():
    from PIL import Image, ImageDraw, ImageFont
    W, H = 620, 1040
    img = Image.new("RGB", (W, H), "#fbfaf6")
    d = ImageDraw.Draw(img)
    mono = "/System/Library/Fonts/Menlo.ttc"
    f = ImageFont.truetype(mono, 22)
    fb = ImageFont.truetype(mono, 28)
    lineas = [
        ("TABERNA EL LUCERNARIO", fb, "c"), ("Restaurante · Cocina de mercado", f, "c"),
        ("C/ Fuencarral 112 · 28010 Madrid", f, "c"), ("CIF B-87654321", f, "c"), ("", f, "l"),
        ("FACTURA SIMPLIFICADA  T-2026/08841", f, "l"), ("Fecha: 26/09/2026   Hora: 15:12", f, "l"),
        ("Mesa: 7      Comensales: 3", f, "l"), ("-" * 40, f, "l"),
        ("3  Menú del día        16,50   49,50", f, "l"), ("1  Vino tinto crianza  18,00   18,00", f, "l"),
        ("2  Agua mineral         2,20    4,40", f, "l"), ("3  Café solo            1,80    5,40", f, "l"),
        ("1  Postre especial      9,10    9,10", f, "l"), ("-" * 40, f, "l"),
        ("Base imponible               78,55", f, "l"), ("IVA 10 %                       7,85", f, "l"),
        ("TOTAL              86,40 €", fb, "l"), ("", f, "l"),
        ("Pago: TARJETA  **** **** **** 4417", f, "l"), ("", f, "l"), ("Gracias por su visita", f, "c")]
    y = 50
    for txt, fo, al in lineas:
        w = d.textlength(txt, font=fo)
        x = (W - w) / 2 if al == "c" else 40
        d.text((x, y), txt, font=fo, fill="#2a2a2a")
        y += 42 if fo is fb else 36
    img = img.rotate(-2.5, expand=True, fillcolor="#d9d4ca")
    assert abs(49.50 + 18.00 + 4.40 + 5.40 + 9.10 - 86.40) < 0.001
    img.save(ruta(GASTOS, "ticket_comida_cliente.jpg"), quality=88)


# ------------------------------------------------------------------ Informe de mercado con fuentes verificadas
def plantilla_verificacion():
    wb = Workbook()
    ws = wb.active
    ws.title = "Verificación"
    filas = [[i, "", "", "", "", "", "", ""] for i in range(1, 6)]
    hoja(ws, ["N.º", "Dato que afirma Claude", "Fuente que cita (URL)", "¿La fuente dice exactamente eso? (sí / no / en parte)",
              "Fecha del dato", "¿Es la fuente original? (sí / no)", "Veredicto (válido / corregir / descartar)", "Notas"],
         filas, [6, 40, 36, 30, 16, 22, 26, 30])
    ws2 = wb.create_sheet("Cómo usarla")
    for t in ["1. Elige los tres datos del informe que más pesan en la decisión.",
              "2. Abre la fuente citada y busca la frase o la cifra exacta.",
              "3. Mira la fecha: un dato correcto de hace cinco años puede no servir hoy.",
              "4. Si la fuente es un artículo que resume otra fuente, busca la original.",
              "5. Anota el veredicto y lleva las correcciones a Claude."]:
        ws2.append([t])
    ws2.column_dimensions["A"].width = 90
    wb.save(ruta(MERCADO, "plantilla_verificacion_fuentes.xlsx"))


# ------------------------------------------------------------------ Prepara tu semana con tu correo y tu calendario
def agenda_bruma():
    ev = [
        ("20261116T093000", "20261116T100000", "Reunión de equipo semanal", "Oficina"),
        ("20261116T120000", "20261116T130000", "Llamada con Hotel Alcázar (boda Ruiz-Ortega)", "Teléfono"),
        ("20261116T160000", "20261116T170000", "Visita a la finca Los Olivos", "Finca Los Olivos"),
        ("20261117T100000", "20261117T113000", "Presentación de la propuesta del Congreso Farmacéutico", "Oficina del cliente"),
        ("20261117T133000", "20261117T143000", "Comida con Catering Sabores", "Restaurante Casa Pepa"),
        ("20261118T090000", "20261118T130000", "Montaje de la gala de la Fundación Mar (fuera de la oficina)", "Palacio de Congresos"),
        ("20261118T170000", "20261118T173000", "Revisión del presupuesto de la boda Ruiz-Ortega", "Oficina"),
        ("20261119T110000", "20261119T120000", "Reunión con Tecnova: lanzamiento de producto", "Videollamada"),
        ("20261119T113000", "20261119T123000", "Reunión con la agencia audiovisual", "Oficina"),
        ("20261119T160000", "20261119T180000", "Preparar dossier Tecnova", "Oficina"),
        ("20261120T100000", "20261120T110000", "Cierre semanal con dirección", "Oficina"),
    ]
    out = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Curso Introduccion a Claude//Bruma Eventos//ES",
           "CALSCALE:GREGORIAN", "X-WR-CALNAME:Bruma Eventos (demo)", "X-WR-TIMEZONE:Europe/Madrid"]
    for i, (a, b, s, loc) in enumerate(ev, 1):
        out += ["BEGIN:VEVENT", f"UID:bruma-demo-{i}@curso-claude.example", "DTSTAMP:20261109T090000Z",
                f"DTSTART;TZID=Europe/Madrid:{a}", f"DTEND;TZID=Europe/Madrid:{b}", f"SUMMARY:{s}",
                f"LOCATION:{loc}", "END:VEVENT"]
    out.append("END:VCALENDAR")
    ruta(SEMANA, "agenda_semana_bruma.ics").write_text("\r\n".join(out) + "\r\n", encoding="utf-8")


# ------------------------------------------------------------------ Crea tu primera skill
CANALES = ["Tienda Centro", "Tienda Playa", "Online"]
CATEGORIAS = {"Bicicletas": (480, 1450), "Accesorios": (12, 85), "Taller": (25, 90), "Alquiler": (15, 45)}


def ventas_rueda_libre():
    # Se busca la primera semilla que cumpla el guion: semanas 45 y 46 estables
    # (ningún canal varía más de un 10 %) y la Tienda Playa cae más de un 20 % en la 47.
    for semilla in range(30, 3000):
        filas, tot = _ventas(semilla)
        estables = all(abs(tot[(s, c)] / tot[(s - 1, c)] - 1) < 0.10 for s in (45, 46) for c in CANALES)
        caida = tot[(47, "Tienda Playa")] / tot[(46, "Tienda Playa")] - 1 < -0.20
        resto = all(abs(tot[(47, c)] / tot[(46, c)] - 1) < 0.12 for c in ("Tienda Centro", "Online"))
        if estables and caida and resto:
            break
    else:
        raise SystemExit("No se encontró una semilla válida para las ventas de Rueda Libre")
    _informe_rueda_libre(filas, tot)


def _ventas(semilla):
    rnd = random.Random(semilla)
    lunes = {44: dt.date(2026, 10, 26), 45: dt.date(2026, 11, 2), 46: dt.date(2026, 11, 9), 47: dt.date(2026, 11, 16)}
    # Multiplicadores por semana y canal: la Tienda Playa cae en la semana 47
    mult = {44: {"Tienda Centro": 1.0, "Tienda Playa": 1.0, "Online": 1.0},
            45: {"Tienda Centro": 1.03, "Tienda Playa": 0.98, "Online": 1.05},
            46: {"Tienda Centro": 1.01, "Tienda Playa": 1.02, "Online": 1.03},
            47: {"Tienda Centro": 1.03, "Tienda Playa": 0.70, "Online": 1.06}}
    base_uds = {"Bicicletas": 2.2, "Accesorios": 9, "Taller": 5, "Alquiler": 6}
    filas = []
    for sem, l in lunes.items():
        for dia in range(6):  # de lunes a sábado
            fecha = l + dt.timedelta(days=dia)
            for canal in CANALES:
                for cat, (lo, hi) in CATEGORIAS.items():
                    if canal == "Online" and cat in ("Taller", "Alquiler"):
                        continue
                    if canal == "Tienda Centro" and cat == "Alquiler":
                        continue
                    uds = max(0, round(rnd.gauss(base_uds[cat], base_uds[cat] * 0.25) * mult[sem][canal]))
                    if uds == 0:
                        continue
                    imp = round(sum(rnd.uniform(lo, hi) for _ in range(uds)), 2)
                    filas.append([sem, fecha, canal, cat, uds, imp])
    tot = {(s, c): round(sum(r[5] for r in filas if r[0] == s and r[2] == c), 2) for s in lunes for c in CANALES}
    return filas, tot


def _informe_rueda_libre(filas, tot):
    lunes = (44, 45, 46, 47)
    for s in (45, 46, 47):
        CONTROL[f"Skill · semana {s}"] = {c: f"{eur(tot[(s, c)])} ({(tot[(s, c)] - tot[(s - 1, c)]) / tot[(s - 1, c)] * 100:+.1f} %)" for c in CANALES}
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"
    hoja(ws, ["Semana", "Fecha", "Canal", "Categoría", "Unidades", "Importe (€)"], filas, [9, 12, 16, 14, 11, 13])
    for r in ws.iter_rows(min_row=2):
        r[1].number_format = "DD/MM/YYYY"
        r[5].number_format = '#,##0.00'
    wb.save(ruta(SKILL, "ventas_rueda_libre_semanas_44-47.xlsx"))

    # Informe de ejemplo de la semana 45, con cifras reales de la hoja
    t45 = sum(tot[(45, c)] for c in CANALES)
    t44 = sum(tot[(44, c)] for c in CANALES)
    cat45 = {k: round(sum(r[5] for r in filas if r[0] == 45 and r[3] == k), 2) for k in CATEGORIAS}
    top = max(cat45, key=cat45.get)
    var = lambda a, b: f"{(a - b) / b * 100:+.1f} %".replace(".", ",")
    d = docx.Document()
    d.styles["Normal"].font.name = "Calibri"
    d.styles["Normal"].font.size = Pt(11)
    d.add_heading("Rueda Libre · Informe semanal de ventas · Semana 45", level=1)
    d.add_paragraph("Del 2 al 7 de noviembre de 2026 · Para: dirección · Preparado por: administración")
    d.add_heading("Resumen", level=2)
    for x in [f"Ventas totales: {eur(t45)} ({var(t45, t44)} frente a la semana 44).",
              f"La categoría que más vende es {top}: {eur(cat45[top])}.",
              f"El canal online es el que más crece: {var(tot[(45, 'Online')], tot[(44, 'Online')])}."]:
        d.add_paragraph(x, style="List Bullet")
    d.add_heading("Ventas por canal", level=2)
    t = d.add_table(rows=1, cols=4)
    t.style = "Light Grid Accent 1"
    for c, v in zip(t.rows[0].cells, ["Canal", "Semana 44", "Semana 45", "Variación"]):
        c.text = v
    for c in CANALES:
        cells = t.add_row().cells
        for cell, v in zip(cells, [c, eur(tot[(44, c)]), eur(tot[(45, c)]), var(tot[(45, c)], tot[(44, c)])]):
            cell.text = v
    cells = t.add_row().cells
    for cell, v in zip(cells, ["Total", eur(t44), eur(t45), var(t45, t44)]):
        cell.text = v
    d.add_heading("Alertas", level=2)
    d.add_paragraph("Sin alertas esta semana: ningún canal cae más de un 15 % respecto a la semana anterior.")
    d.add_heading("Para la próxima semana", level=2)
    d.add_paragraph("Revisar el stock de bicicletas eléctricas antes de la campaña de Black Friday.")
    d.save(ruta(SKILL, "informe_ejemplo_semana_45.docx"))


# ------------------------------------------------------------------ Delega una tarea completa en Cowork
EMISOR = "Taller Arcilla · C/ Alfareros 7 · 46001 Valencia · NIF B-46123987"
FACTURAS = [
    # (archivo, proveedor, nif, número, fecha, [(concepto, cantidad, precio)], iva)
    (["scan_0007.pdf", "factura ceramicas sur.pdf"], "Cerámicas del Sur SL", "B-41556677", "FS-2026-1184", "08/10/2026",
     [("Arcilla roja (saco 25 kg)", 12, 28.50), ("Arcilla blanca (saco 25 kg)", 6, 24.00)], 21),
    (["IMG_2041.pdf"], "Esmaltes Iris", "B-46998811", "EI/26/0457", "03/10/2026",
     [("Esmalte transparente 5 kg", 4, 46.10), ("Óxido de cobalto 500 g", 2, 64.00)], 21),
    (["Factura_EC-889213.pdf"], "Energía Clara SA", "A-28445566", "EC-889213", "31/10/2026",
     [("Consumo eléctrico octubre (horno incluido)", 1, 418.77)], 21),
    (["doc20261015.pdf"], "Hornos Ceramtec", "B-08776655", "HC-2026-077", "15/10/2026",
     [("Mantenimiento anual del horno", 1, 260.00)], 21),
    (["ticket papeleria.pdf"], "Papelería Trazo", "B-46332211", "PT-5521", "06/10/2026",
     [("Etiquetas adhesivas", 3, 6.45), ("Cuadernos de pedidos", 4, 7.25)], 21),
    (["FS-2026-1231 (1).pdf"], "Cerámicas del Sur SL", "B-41556677", "FS-2026-1231", "22/10/2026",
     [("Arcilla refractaria (saco 25 kg)", 8, 30.375)], 21),
    (["transporte.pdf"], "Transportes Rápido Levante", "B-46557788", "TRL-26-3390", "09/10/2026",
     [("Envío de pedido a Madrid (3 bultos)", 1, 75.00)], 21),
    (["factura(1).pdf"], "Telecom Viva", "A-78112233", "TV-2026-10-55821", "01/10/2026",
     [("Fibra y teléfono, octubre", 1, 49.50)], 21),
    (["alquiler oct.pdf"], "Inmobiliaria Portal Norte", "B-46221100", "PN-10/2026", "01/10/2026",
     [("Alquiler del local, octubre", 1, 950.00)], 21),
    (["AF-2026-0312.pdf"], "Asesoría Ferrer", "B-46889900", "AF-2026-0312", "30/10/2026",
     [("Asesoría fiscal y laboral, octubre", 1, 180.00)], 21),
    (["scan_0011.pdf"], "Embalajes Caja Fuerte", "B-46443322", "ECF-7731", "14/10/2026",
     [("Cajas de cartón doble canal", 60, 1.36), ("Papel burbuja (rollo)", 4, 12.75)], 21),
    (["esmaltes 2.pdf"], "Esmaltes Iris", "B-46998811", "EI/26/0489", "20/10/2026",
     [("Esmalte mate blanco 2 kg", 4, 24.20)], 21),
    (["poliza_seguro.pdf"], "Seguros Amparo", "A-28990011", "SA-26-10-4471", "05/10/2026",
     [("Seguro multirriesgo del taller, cuota de octubre (operación exenta de IVA)", 1, 64.30)], 0),
    (["libro.pdf"], "Librería Técnica Arcadia", "B-46667788", "LTA-2290", "17/10/2026",
     [("Libro «Esmaltes de alta temperatura»", 1, 38.46)], 4),
    (["factura(2).pdf"], "Limpiezas Brillo", "B-46119988", "LB-2026-1010", "31/10/2026",
     [("Limpieza del taller, octubre", 1, 120.00)], 21),
    (["imprenta_final_FINAL.pdf"], "Imprenta Gutenberg", "B-46776655", "IG-26-812", "27/10/2026",
     [("Tarjetas de visita (500 uds.)", 1, 45.00), ("Folletos del curso de torno (1.000 uds.)", 1, 110.00)], 21),
]
CSS_FACTURA = """
@page { size: A4; margin: 18mm; }
body { font-family: 'Helvetica Neue', Arial, sans-serif; font-size: 10.5pt; color: #222; }
.top { display: flex; justify-content: space-between; border-bottom: 2px solid #333; padding-bottom: 4mm; }
.top h1 { font-size: 20pt; margin: 0; }
.datos { margin: 6mm 0; display: flex; gap: 20mm; }
table { width: 100%; border-collapse: collapse; margin-top: 4mm; }
th, td { border-bottom: 1px solid #ccc; padding: 2mm; text-align: right; }
th:first-child, td:first-child { text-align: left; }
.tot { width: 70mm; margin-left: auto; margin-top: 6mm; }
.tot td { border: none; }
.tot tr:last-child td { font-weight: bold; font-size: 12pt; border-top: 2px solid #333; }
.pie { margin-top: 20mm; font-size: 8.5pt; color: #777; }
"""


def facturas_arcilla():
    base_dir = MATERIALES / carpeta(COWORK) / "Facturas_proveedores_octubre"
    if base_dir.exists():
        shutil.rmtree(base_dir)
    total_general = 0
    for archivos, prov, nif, num, fecha, lineas, iva in FACTURAS:
        base = round(sum(c * p for _, c, p in lineas), 2)
        cuota = round(base * iva / 100, 2)
        total = round(base + cuota, 2)
        total_general += total
        filas = "".join(f"<tr><td>{c}</td><td>{q}</td><td>{eur(p)}</td><td>{eur(q * p)}</td></tr>" for c, q, p in lineas)
        iva_txt = f"IVA {iva} %" if iva else "IVA (exento)"
        html = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><style>{CSS_FACTURA}</style></head><body>
<div class="top"><div><h1>{prov}</h1><div>NIF {nif}</div></div><div style="text-align:right"><b>FACTURA</b><br>N.º {num}<br>Fecha: {fecha}</div></div>
<div class="datos"><div><b>Cliente</b><br>{EMISOR.replace(' · ', '<br>')}</div><div><b>Forma de pago</b><br>Transferencia a 30 días</div></div>
<table><tr><th>Concepto</th><th>Cantidad</th><th>Precio</th><th>Importe</th></tr>{filas}</table>
<table class="tot"><tr><td>Base imponible</td><td>{eur(base)}</td></tr><tr><td>{iva_txt}</td><td>{eur(cuota)}</td></tr>
<tr><td>Total</td><td>{eur(total)}</td></tr></table>
<div class="pie">Documento ficticio para uso formativo.</div></body></html>"""
        for a in archivos:
            html_a_pdf(html, base_dir / a)
    pres = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><style>{CSS_FACTURA}</style></head><body>
<div class="top"><div><h1>Reformas Albañil Pérez</h1><div>NIF 22334455-K</div></div><div style="text-align:right"><b>PRESUPUESTO</b><br>N.º P-2026-41<br>Fecha: 12/10/2026</div></div>
<div class="datos"><div><b>Cliente</b><br>{EMISOR.replace(' · ', '<br>')}</div><div><b>Validez</b><br>30 días</div></div>
<table><tr><th>Concepto</th><th>Importe</th></tr><tr><td>Reforma del baño del taller (alicatado, plato de ducha y fontanería)</td><td>{eur(2340)}</td></tr></table>
<p>Este documento es un presupuesto y no tiene validez como factura. IVA no incluido.</p>
<div class="pie">Documento ficticio para uso formativo.</div></body></html>"""
    html_a_pdf(pres, base_dir / "doc20261012.pdf")
    CONTROL["Cowork · facturas distintas"] = len(FACTURAS)
    CONTROL["Cowork · archivos en la carpeta"] = sum(len(f[0]) for f in FACTURAS) + 1
    CONTROL["Cowork · total de las facturas (sin duplicado)"] = eur(total_general)


def main():
    textos()
    informe_arrabal()
    gastos_lumbre()
    ticket_restaurante()
    plantilla_verificacion()
    agenda_bruma()
    ventas_rueda_libre()
    facturas_arcilla()
    print("Cifras de control:")
    for k, v in CONTROL.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
