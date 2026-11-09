"""Genera un PDF de enunciado por clase práctica en '1-Enunciados/'.

Uso: python3 -I .generadores/generar_enunciados.py   (macOS, usa Google Chrome headless)
Misma maqueta que el curso Claude Associate Foundations: diapositivas 16:9
(portada, presentación del caso y una diapositiva por paso). Cada diapositiva
se reduce sola si su contenido no cabe; el script avisa si baja del 85 %.
La Clase 05 (crear la cuenta) no tiene enunciado.
"""
import html
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from comun import CHROME, CLASES, ENUNCIADOS  # noqa: E402


class Bloque(str):
    """Texto que no se escribe en el chat (instrucciones de un proyecto, una
    descripción de skill…). Se pinta en otro color con su etiqueta encima."""
    def __new__(cls, texto, etiqueta="Instrucciones del proyecto"):
        obj = super().__new__(cls, texto)
        obj.etiqueta = etiqueta
        return obj


# Cada paso: (título, concepto, qué haces, [prompts o Bloque], [lista opcional]).
EJERCICIOS = [
  {
    "n": 7,
    "titulo": "Una clienta molesta y un correo que hay que acertar",
    "caso": "Una tienda online de cosmética recibe la queja de una clienta fiel. Se pasa de una petición de una línea a un correo listo para enviar.",
    "archivos": "queja_clienta.txt y correo_ejemplo_marca.txt",
    "modelo": "El modelo por defecto en todos los pasos: así la mejora se debe a la petición, no al modelo.",
    "pasos": [
      ("La petición vaga", "Una persona nueva muy brillante",
       "Pides el correo sin contar nada. Guarda la respuesta para compararla al final.",
       ["Escríbeme un correo para una clienta que se queja."]),
      ("Contexto, tarea y formato", "Los tres ingredientes y el porqué",
       "Dices quién eres, qué ha pasado, qué quieres y cómo lo quieres. Y explicas por qué.",
       ["Soy responsable de atención al cliente en Lúa, una tienda online de cosmética natural.\n"
        "Una clienta habitual ha recibido tarde un pack regalo y con un frasco roto. Te pego su correo y nuestras notas internas.\n"
        "Escríbele una respuesta que pida disculpas, explique qué ha pasado y le ofrezca las dos soluciones que permitimos.\n"
        "Máximo 150 palabras, en tono cercano y tuteando, porque es una clienta fiel y queremos que siga comprando.\n"
        "No prometas nada que no esté en las notas.\n"
        "[pega queja_clienta.txt]"]),
      ("Un ejemplo vale más que una explicación", "Ejemplos, tono y estructura",
       "Le das una respuesta real del equipo para que copie el estilo, no las frases.",
       ["Así respondemos normalmente en Lúa. Reescribe el correo siguiendo el tono, la estructura y la firma de este ejemplo, sin copiar sus frases.\n"
        "[pega correo_ejemplo_marca.txt]"]),
      ("Afinar sin empezar de cero", "Indicaciones concretas y autorrevisión",
       "Corriges con instrucciones concretas y le pides que revise el correo con los ojos de la clienta.",
       ["Más corto: quita la explicación del transportista y pon las dos opciones en una lista.",
        "Lee el correo como lo leería esta clienta, que está molesta porque el regalo no llegó a tiempo. "
        "Dime qué le podría sentar mal y qué mejorarías antes de que lo envíe."],
       ["Al terminar, pon lado a lado el correo del paso 1 y el final."]),
    ],
  },
  {
    "n": 9,
    "titulo": "Un club deportivo, dos tareas y tres modelos",
    "caso": "Un club deportivo clasifica los mensajes de sus socios y elige un software de reservas. La misma tarea con distintos modelos y niveles de esfuerzo.",
    "archivos": "mensajes_socios.txt, clave_mensajes_socios.txt y ofertas_software_reservas.txt",
    "modelo": "Haiku, Sonnet y Opus, cada prueba en una conversación nueva. Si tu plan no incluye alguno, compara los que tengas.",
    "pasos": [
      ("Tarea simple: ligero frente a potente", "El modelo mínimo que cumple",
       "Fijas el criterio antes de probar: 18 de 20 aciertos y una tabla. Lanzas el mismo prompt con Haiku y con Opus, en conversaciones separadas, y corriges con la clave.",
       ["Clasifica cada mensaje en una de estas categorías: altas y bajas, reservas, facturación, quejas o información.\n"
        "Devuelve solo una tabla id | categoría.\n"
        "[pega mensajes_socios.txt]"]),
      ("El esfuerzo", "Más esfuerzo no siempre es mejor",
       "Repites la misma clasificación con Sonnet en esfuerzo bajo y en esfuerzo alto. Compara el resultado y el tiempo que tarda.",
       [],
       ["Si los dos cumplen el criterio, gana el esfuerzo bajo: tarda menos y gasta menos de tu plan.",
        "El esfuerzo se elige en el mismo menú que el modelo, junto al botón de enviar."]),
      ("Tarea compleja: varios factores", "Cuándo subir de modelo",
       "Una decisión con condiciones, costes y una letra pequeña. Lanzas el mismo prompt con Haiku y con Opus en esfuerzo alto.",
       ["Con estas tres ofertas y nuestras necesidades, recomienda qué software de reservas debería contratar el Club Deportivo Marea.\n"
        "Calcula el coste del primer y del segundo año de cada opción.\n"
        "Razona antes de concluir, indica qué supuestos haces y qué dato cambiaría tu recomendación.\n"
        "[pega ofertas_software_reservas.txt]"],
       ["Comprueba si cada modelo detecta el precio promocional de B y el TPV de C."]),
      ("Mala petición, mejor modelo", "La petición pesa más que el modelo",
       "Opus con una petición vaga y solo las ofertas, frente a Sonnet con el prompt completo del paso anterior.",
       ["¿Cuál de estos programas es mejor?\n[pega solo la parte de OFERTAS RECIBIDAS]"],
       ["Repite en Sonnet el prompt del paso 3 y compara las dos respuestas."]),
    ],
  },
  {
    "n": 12,
    "titulo": "El informe anual de un obrador de pan",
    "caso": "Un informe de 7 páginas con una trampa: el texto dice que las mermas están bajo control y el gráfico dice lo contrario. Y hay un dato que no aparece.",
    "archivos": "informe_anual_2026_panes_arrabal.pdf",
    "modelo": "El modelo por defecto, el mismo en todos los pasos.",
    "pasos": [
      ("Resumir sin más", "Lo que pasa si no dices para qué",
       "Adjuntas el PDF y pides un resumen sin ninguna indicación.",
       ["Resume este informe."]),
      ("Un informe ejecutivo de verdad", "Contexto, audiencia y formato",
       "Dices quién lo va a leer y para qué, fijas el formato y le pides que indique de dónde sale cada dato.",
       ["Soy la nueva responsable de operaciones de Panes Arrabal.\n"
        "Prepara un informe ejecutivo de una página para el comité de dirección del lunes: 5 puntos clave con cifras, 3 riesgos y 2 decisiones que habría que tomar.\n"
        "Usa solo lo que dice el informe e indica el apartado del que sale cada dato."]),
      ("Preguntas que dan respuestas fiables", "Citas, solo el documento y «no lo dice»",
       "Preguntas concretas, con el texto literal como prueba. La tercera pregunta no tiene respuesta en el informe.",
       ["Responde a estas preguntas usando solo el informe.\n"
        "Para cada respuesta, cita el texto literal y el apartado. Si el informe no lo dice, escribe «El informe no lo dice».\n"
        "1. ¿Qué tienda vendió menos que el año anterior y qué explicación da el informe?\n"
        "2. ¿Cuándo vence el contrato de alquiler de la tienda Centro?\n"
        "3. ¿Cuál fue el beneficio neto de 2026?"]),
      ("Lo que está en un gráfico", "Comprobar lo que no está en el texto",
       "Preguntas directamente por el gráfico del apartado 4 y lo comparas con lo que dice el texto. Después actualizas el informe ejecutivo.",
       ["¿Qué muestra el gráfico de mermas del apartado 4? ¿Coincide con lo que dice el texto de ese apartado?",
        "Actualiza el informe ejecutivo con lo que has encontrado en el gráfico."]),
    ],
  },
  {
    "n": 14,
    "titulo": "Los gastos del trimestre de un estudio de arquitectura",
    "caso": "Una hoja con 75 gastos del tercer trimestre. Tiene un cargo duplicado, categorías mal escritas y una devolución. Además, falta un ticket que solo está en una foto.",
    "archivos": "gastos_T3_2026_estudio_lumbre.xlsx y ticket_comida_cliente.jpg",
    "modelo": "El modelo por defecto, con la opción de crear y editar archivos activada.",
    "pasos": [
      ("Preguntar en lenguaje normal", "Análisis de datos sin fórmulas",
       "Adjuntas la hoja y preguntas lo que quieres saber, sin fórmulas.",
       ["Te adjunto los gastos del tercer trimestre de nuestro estudio.\n"
        "Dime el total por categoría y por mes, y qué partida ha crecido más de julio a septiembre y por qué.\n"
        "Devuélvelo en una tabla y tres conclusiones."]),
      ("Comprobar el análisis", "Supuestos, totales conocidos y filas",
       "Le das un total que conoces y le pides que explique lo que ha supuesto. Si no cuadra, que te diga de qué filas sale.",
       ["El extracto del banco suma 15.407,60 € para el trimestre. ¿Coincide con tu total?\n"
        "Explica qué has supuesto sobre los datos: categorías, importes negativos y posibles duplicados.\n"
        "Si hay diferencias, dime de qué filas salen."]),
      ("Un gasto que está en una foto", "Claude también ve",
       "Adjuntas la foto del ticket y pides que lo convierta en una fila de la hoja.",
       ["Extrae de esta foto la fecha, el proveedor, el importe total y el IVA, y añádelo como una fila nueva con el formato de la hoja.\n"
        "Si algún dato no se lee bien, dímelo en lugar de suponerlo.\n"
        "[adjunta ticket_comida_cliente.jpg]"]),
      ("Un archivo que puedes usar", "Crear un Excel con fórmulas y gráfico",
       "Pides una hoja nueva lista para trabajar y compruebas una fórmula al abrirla.",
       ["Devuélveme un Excel nuevo con tres pestañas:\n"
        "1. Los gastos corregidos, con el ticket añadido y sin el duplicado.\n"
        "2. Un resumen por categoría y mes, con fórmulas.\n"
        "3. Un gráfico de barras del gasto por mes."],
       ["Abre el archivo y haz clic en un total del resumen: debe ser una fórmula, no un número escrito."]),
    ],
  },
  {
    "n": 16,
    "titulo": "La guía de bienvenida de una cadena de tiendas",
    "caso": "Unas notas desordenadas para las nuevas incorporaciones de Ludo, tres tiendas de juegos de mesa, se convierten en una guía visual. Entre las notas hay datos que nunca deberían compartirse.",
    "archivos": "notas_bienvenida_ludo.txt",
    "modelo": "El modelo por defecto.",
    "pasos": [
      ("De unas notas a una guía", "Qué es un artifact",
       "Pegas las notas tal cual y pides una guía para leer en el móvil.",
       ["Con estas notas, crea una guía de bienvenida visual para las personas que se incorporan a nuestras tiendas.\n"
        "Que se lea en el móvil en cinco minutos: secciones cortas, iconos y lo más importante arriba.\n"
        "[pega notas_bienvenida_ludo.txt]"]),
      ("Cambios concretos", "Afinar un artifact",
       "Pides cambios concretos, uno detrás de otro, sobre el mismo artifact.",
       ["Usa los colores de nuestra marca: #1F4E5F y #F2A541. Haz los títulos más grandes.",
        "Añade al final una lista de comprobación interactiva para la primera semana, día por día."]),
      ("Un diagrama de proceso", "Gráficos y diagramas",
       "Añades un diagrama para el proceso que más dudas genera.",
       ["Añade un diagrama sencillo con los pasos para gestionar una devolución, desde que llega el cliente hasta que se registra en caja. "
        "Incluye qué pasa si no trae ticket o si el juego está desprecintado."]),
      ("Revisar antes de compartir", "Compartir con control",
       "Antes de enviar el enlace, revisas qué datos contiene. Después lo compartes solo para ver.",
       ["Antes de compartir esta guía: revisa si contiene datos que no deberían salir del equipo, como contraseñas, códigos o teléfonos personales. "
        "Dime cuáles son, quítalos y sustitúyelos por dónde consultarlos."],
       ["Compártelo con enlace y con permiso solo para ver."]),
    ],
  },
  {
    "n": 18,
    "titulo": "¿Abrimos un coworking en Málaga?",
    "caso": "Una socia tiene que decidir si invierte en un espacio de coworking. Se pasa de una búsqueda rápida a una investigación a fondo y se comprueban las fuentes una a una.",
    "archivos": "plantilla_verificacion_fuentes.xlsx",
    "modelo": "El modelo por defecto, con la búsqueda web activada. El paso 2 usa el modo de investigación.",
    "pasos": [
      ("Una búsqueda rápida", "Búsqueda web con fuentes",
       "Una pregunta concreta que depende de datos actuales.",
       ["¿Cuántos espacios de coworking hay en Málaga capital y cuál es el precio medio de un puesto fijo al mes?\n"
        "Indica la fuente y la fecha de cada dato."]),
      ("La investigación a fondo", "Pedir bien una investigación",
       "Activas el modo de investigación con una petición completa: para qué es, quién lo lee, periodo, fuentes y qué hacer si falta un dato.",
       ["Necesito un informe sobre el mercado de coworking en Málaga capital para decidir si abrimos un espacio de 40 puestos en 2027.\n"
        "Lo leerá una socia que tiene que aprobar la inversión.\n"
        "Céntrate en los dos últimos años y en fuentes oficiales, medios económicos y las webs de los propios espacios.\n"
        "Incluye la oferta actual y sus precios, el perfil de la demanda, las tendencias y los riesgos.\n"
        "Separa los datos comprobados de tus conclusiones y, si no encuentras una cifra, indícalo en lugar de estimarla."]),
      ("Comprobar tres fuentes", "Citar no es verificar",
       "Eliges los tres datos que más pesan en la decisión y los compruebas con la plantilla.",
       [],
       ["Abre la fuente y busca la frase o la cifra exacta.",
        "Mira la fecha: un dato correcto de hace cinco años puede no servir hoy.",
        "Prioriza la fuente original frente al artículo que la resume.",
        "Anota el veredicto de cada dato en plantilla_verificacion_fuentes.xlsx."]),
      ("Corregir con lo que has comprobado", "Separar hechos y conclusiones",
       "Le devuelves lo que has encontrado y le pides que marque lo que no está verificado.",
       ["He revisado tres datos del informe:\n[pega tus veredictos de la plantilla]\n"
        "Corrige el informe con lo que he encontrado y marca con [SIN VERIFICAR] cualquier cifra que se apoye en una sola fuente secundaria."]),
    ],
  },
  {
    "n": 20,
    "titulo": "Renovar el contrato con un cliente que quiere recortar",
    "caso": "Una agencia de diseño web tiene que renovar el contrato de mantenimiento de una cadena hotelera con un nuevo director financiero que pide recortar un 10 %.",
    "archivos": "datos_renovacion_mar_blau.txt",
    "modelo": "El modelo por defecto, toda la demo en la misma conversación.",
    "pasos": [
      ("Que Claude te entreviste", "Darle la vuelta al proceso",
       "En lugar de explicarlo todo, dejas que Claude pregunte. Respondes con los datos del archivo, sin pegarlos de golpe.",
       ["Quiero escribir una propuesta para renovar el contrato de mantenimiento web con un cliente.\n"
        "Antes de redactar nada, hazme las preguntas que necesites, de una en una, hasta que tengas claro lo que necesito."]),
      ("El borrador", "Formato y siguiente paso",
       "Con el contexto recogido, pides el borrador con una estructura clara.",
       ["Con todo lo que te he contado, escribe la propuesta.\n"
        "Máximo una página. Empieza por los resultados de este año, presenta las dos opciones en una tabla y termina con el siguiente paso y la fecha."]),
      ("El lector más exigente", "Una perspectiva concreta",
       "Pides la crítica desde la mirada de quien tiene que decir que sí.",
       ["Lee esta propuesta como el nuevo director financiero de Hoteles Mar Blau, que quiere recortar un 10 % el gasto en proveedores y desconfía de las renovaciones con subida.\n"
        "Dime qué objeciones pondría, qué le parecería poco claro y qué le faltaría para decir que sí."]),
      ("Mejorar con criterio", "La crítica es una opinión",
       "Eliges qué objeciones aplicar y cuáles no, y explicas por qué. Después compruebas las cifras.",
       ["Aplica las objeciones [números que aceptas]. La [número] no, porque [tu motivo].\n"
        "Reescribe la propuesta con esos cambios y sin pasar de una página."],
       ["Comprueba que los precios, los porcentajes y la fecha límite coinciden con el archivo."]),
    ],
  },
  {
    "n": 22,
    "titulo": "Decirle a alguien que no ha conseguido el puesto",
    "caso": "La responsable de un equipo de atención al cliente tiene que contar a una agente sénior que la plaza de coordinación ha sido para otra persona. Se prepara, se ensaya y se escribe el correo de seguimiento.",
    "archivos": "contexto_conversacion.txt y mis_correos_ejemplo.txt",
    "modelo": "El modelo por defecto. El ensayo del paso 2 se puede hacer con el modo de voz de la aplicación del móvil.",
    "pasos": [
      ("Preparar la conversación", "Pensar con Claude",
       "Le das el contexto y le pides una estructura para la conversación.",
       ["Tengo que tener una conversación difícil y quiero prepararla. Te paso el contexto.\n"
        "Ayúdame con el objetivo de la conversación, los tres mensajes clave, cómo empezar y las reacciones más probables de Marta.\n"
        "[pega contexto_conversacion.txt]"]),
      ("El ensayo", "Claude hace de la otra persona",
       "Claude hace de Marta. Respondes como lo harías de verdad, un mensaje cada vez.",
       ["Ahora haz de Marta. Reacciona como una persona decepcionada que siente que no se ha valorado su trabajo: no me lo pongas fácil, pero sé realista.\n"
        "Responde solo como Marta, un mensaje cada vez, y espera mi respuesta.\n"
        "Empiezo yo: «Marta, gracias por venir. Quería hablar contigo de la plaza de coordinación.»"]),
      ("El feedback", "Discernir: juzgar el resultado",
       "Le pides que salga del papel y te evalúe con criterios concretos.",
       ["Sal del papel. Evalúa cómo lo he hecho según estos criterios: claridad del mensaje, empatía, si he prometido algo que no puedo cumplir y si he cerrado con próximos pasos.\n"
        "Dame dos frases concretas que podría haber dicho mejor."]),
      ("El correo de seguimiento, con tu estilo", "Enseñarle cómo escribes",
       "Primero describe tu estilo a partir de tus correos. Después escribe el seguimiento con esas pautas.",
       ["Analiza cómo escribo en estos correos: el tono, la longitud de las frases, cómo saludo y me despido y qué palabras nunca uso. "
        "Descríbelo en una lista de pautas.\n[pega mis_correos_ejemplo.txt]",
        "Siguiendo esas pautas, escribe el correo de seguimiento a Marta con un resumen de lo hablado y los próximos pasos. Máximo 120 palabras."],
       ["Revisa las pautas antes de usarlas: a veces exagera algún rasgo."]),
    ],
  },
  {
    "n": 24,
    "titulo": "El asistente de la oficina de una autoescuela",
    "caso": "Un proyecto con instrucciones y documentos para responder a los alumnos. Funciona bien hasta que alguien sube las tarifas antiguas.",
    "archivos": "Tarifas_y_condiciones_sept2026.docx, Tarifas_y_condiciones.docx (la antigua) y Preguntas_frecuentes.docx",
    "modelo": "El modelo por defecto, dentro del proyecto.",
    "pasos": [
      ("Crear el proyecto", "Instrucciones del proyecto",
       "Creas el proyecto «Oficina Autoescuela Ronda» con estas instrucciones y, sin subir documentos, haces una pregunta de un alumno.",
       [Bloque("Eres el asistente de la oficina de Autoescuela Ronda. Ayudas a responder a los alumnos por WhatsApp y por correo.\n"
               "Responde en un máximo de 80 palabras, tuteando y en tono cercano.\n"
               "Basa los precios y las condiciones solo en los documentos del proyecto y di de cuál sale cada dato.\n"
               "Si algo no está en los documentos, dilo y propón consultarlo con la dirección.\n"
               "Nunca confirmes fechas de examen: las asigna la DGT.", "Instrucciones del proyecto «Oficina Autoescuela Ronda»"),
        "Un alumno pregunta cuánto cuesta el pack de 10 clases y qué pasa si cancela una clase el mismo día."]),
      ("Subir el conocimiento", "El qué en los documentos, el cómo en las instrucciones",
       "Subes solo las tarifas vigentes y las preguntas frecuentes. Repites la pregunta en un chat nuevo del proyecto.",
       ["Un alumno pregunta cuánto cuesta el pack de 10 clases y qué pasa si cancela una clase el mismo día."]),
      ("Romperlo", "Versiones repetidas",
       "Subes también Tarifas_y_condiciones.docx, la versión antigua sin fecha, y repites la pregunta. Después la retiras y vuelves a preguntar.",
       ["Un alumno pregunta cuánto cuesta el pack de 10 clases y qué pasa si cancela una clase el mismo día."],
       ["Un proyecto necesita mantenimiento: cuando un documento cambia, sustitúyelo."]),
      ("Memoria e instrucciones", "Qué va en cada sitio",
       "Le pides que recuerde una preferencia y revisas en los ajustes dónde ha quedado guardada.",
       ["Recuerda que firmo todos los mensajes como «Equipo de Autoescuela Ronda»."],
       ["Instrucciones del perfil: cómo eres y cómo quieres que te responda, en todas las conversaciones.",
        "Instrucciones del proyecto: cómo trabajar en este tema.",
        "Conocimiento del proyecto: los documentos de referencia.",
        "Memoria: lo que Claude aprende de ti. Revísala en los ajustes; lo estable, mejor en las instrucciones."]),
    ],
  },
  {
    "n": 26,
    "titulo": "Una semana de trabajo en una agencia de eventos",
    "caso": "Claude conectado al correo y al calendario prepara la semana: reuniones, solapamientos, correos urgentes y borradores. Uno de los correos esconde una instrucción.",
    "archivos": "agenda_semana_bruma.ics y correos_demo_bruma.txt (para montar una cuenta de pruebas)",
    "modelo": "El modelo por defecto, con los conectores de Gmail y Google Calendar.",
    "pasos": [
      ("Conectar y poner reglas", "Leer no es lo mismo que actuar",
       "Conectas Gmail y Google Calendar desde el directorio de conectores y fijas cómo vas a trabajar.",
       [],
       ["Mejor con una cuenta de pruebas: importa agenda_semana_bruma.ics y envíale los correos del archivo.",
        "Claude solo ve lo que tú puedes ver en esas aplicaciones.",
        "Al principio, borradores y no envíos.",
        "Lee cada petición de aprobación antes de aceptarla."]),
      ("Tu semana en un resumen", "Correo y calendario a la vez",
       "Pides una visión de la semana que cruce agenda y correo.",
       ["Revisa mi calendario de esta semana y los correos de los últimos tres días.\n"
        "Prepárame un resumen con las reuniones de cada día, los conflictos de horario, los correos que necesitan respuesta antes del miércoles y los huecos de al menos una hora para trabajo concentrado."]),
      ("Borradores y un evento", "La aprobación antes de actuar",
       "Claude prepara las respuestas sin enviarlas y propone mover la reunión que se solapa.",
       ["Prepara borradores de respuesta para los correos urgentes, pero no envíes nada.",
        "Propón una hora nueva para la reunión con Tecnova que encaje con lo que piden en su correo. Cuando te la confirme, cambia el evento."]),
      ("La instrucción escondida", "Inyección de instrucciones",
       "Pides el resumen de un correo que intenta darle órdenes a Claude.",
       ["Resume el correo de «Facturación Proveedores» y dime qué me piden."],
       ["Fíjate en si Claude te avisa de la instrucción escondida y en que no reenvía nada.",
        "Conecta solo lo que vas a usar y usa «permitir siempre» solo con lo que conoces bien."]),
    ],
  },
  {
    "n": 28,
    "titulo": "Elegir herramienta de gestión de proyectos",
    "caso": "Un equipo de 8 personas compara tres herramientas en sus webs oficiales. Claude navega, recoge los datos y deja la decisión preparada.",
    "archivos": "criterios_comparativa.txt",
    "modelo": "Claude en Chrome, en un perfil del navegador separado.",
    "pasos": [
      ("Preparar el navegador", "Cómo controlas lo que hace",
       "Antes de empezar, preparas un entorno seguro y decides cuánto control quieres.",
       [],
       ["Un perfil de Chrome separado, sin sesiones abiertas de banco, correo personal o salud.",
        "Permisos por web: una vez, siempre o nunca. Revísalos cuando quieras.",
        "Empieza con tareas de solo lectura: buscar, comparar y resumir.",
        "Elige si apruebas cada acción o el plan completo."]),
      ("Comparar en varias pestañas", "Recopilar datos de varias webs",
       "Claude abre las páginas de precios y rellena la comparativa con los criterios del equipo.",
       ["Abre las páginas de precios de Trello, Asana y monday.com.\n"
        "Para un equipo de 8 personas con facturación anual, recoge en una tabla: el plan más barato que tenga vista de calendario, el precio total al año para las 8 personas, los límites importantes del plan y si hay prueba gratuita.\n"
        "Indica la dirección de la página de la que sale cada dato. No inicies sesión ni rellenes ningún formulario."]),
      ("Comprobar un dato en la web", "Por usuario o en total",
       "Abres tú una de las páginas y compruebas el dato más fácil de malinterpretar.",
       ["¿De qué parte exacta de la página sale el precio de Asana? ¿Es por usuario y mes o el total? "
        "¿Hay un mínimo de usuarios que cambie el precio para 8 personas?"]),
      ("Lo que no hará y lo que puede repetir", "Límites y tareas programadas",
       "Le pides algo que no debe hacer y después programas la revisión de precios.",
       ["Regístrate en la prueba gratuita de la herramienta que hayas recomendado.",
        "Repite esta comparativa el primer lunes de cada mes y avísame solo si cambia algún precio."],
       ["Claude no crea cuentas, no compra y no sigue instrucciones que aparezcan dentro de una web."]),
    ],
  },
  {
    "n": 30,
    "titulo": "El informe de ventas de los lunes",
    "caso": "Cada lunes, una tienda de bicicletas prepara el mismo informe para dirección. Se convierte en una skill que Claude usa sola cuando aparece la tarea.",
    "archivos": "ventas_rueda_libre_semanas_44-47.xlsx e informe_ejemplo_semana_45.docx",
    "modelo": "El modelo por defecto, con la opción de crear y editar archivos activada.",
    "pasos": [
      ("Crear la skill con Claude", "Que Claude te pregunte cómo lo haces",
       "Describes la tarea, adjuntas un informe de ejemplo y respondes a sus preguntas.",
       ["Quiero crear una skill para preparar el informe semanal de ventas que envío cada lunes a dirección.\n"
        "Te adjunto un informe de ejemplo y la hoja de ventas.\n"
        "Hazme las preguntas que necesites, de una en una, sobre cómo lo hago, y después prepara la skill.\n"
        "[adjunta informe_ejemplo_semana_45.docx y ventas_rueda_libre_semanas_44-47.xlsx]"]),
      ("La descripción decide", "Cómo sabe Claude cuándo usarla",
       "Revisas la descripción de la skill: es lo único que Claude lee para decidir si la usa.",
       [Bloque("Ayuda con informes.", "Descripción vaga"),
        Bloque("Prepara el informe semanal de ventas de Rueda Libre para dirección a partir de la hoja de ventas de los lunes: resumen, ventas por canal frente a la semana anterior, alertas y próximos pasos.",
               "Descripción concreta"),
        "Enséñame la descripción de la skill y reescríbela para que diga exactamente cuándo hay que usarla."]),
      ("Probarla en una conversación nueva", "Comprobar que se activa",
       "Guardas la skill, abres una conversación nueva y pides el informe sin mencionarla.",
       ["Prepara el informe de ventas de la semana 46.\n[adjunta ventas_rueda_libre_semanas_44-47.xlsx]"],
       ["Comprueba que Claude dice que está usando la skill y que el informe sigue el formato del ejemplo."]),
      ("Corregir y mejorar", "Una skill se prueba y se afina",
       "Añades una regla que faltaba y vuelves a probar con la semana siguiente.",
       ["Añade a la skill esta regla: si un canal cae más de un 15 % respecto a la semana anterior, la alerta va la primera, en negrita y con una posible causa a revisar.",
        "Prepara el informe de ventas de la semana 47."]),
    ],
  },
  {
    "n": 32,
    "titulo": "Una carpeta de facturas que nadie ha ordenado",
    "caso": "Las facturas de octubre de un taller de cerámica llegan con nombres sin sentido. Hay una duplicada y un documento que no es una factura. Cowork las ordena y prepara el resumen.",
    "archivos": "Facturas_proveedores_octubre/ (18 archivos PDF)",
    "modelo": "Cowork en la aplicación de escritorio, con un modelo potente: va a tomar decisiones sobre tus archivos.",
    "pasos": [
      ("Preparar el terreno", "Cómo supervisar a un agente",
       "Antes de delegar, limitas lo que Claude puede tocar.",
       [],
       ["Trabaja con una copia de la carpeta, nunca con la original.",
        "Dale acceso solo a esa carpeta, no a todo tu ordenador.",
        "Elige si apruebas cada acción o si trabaja solo con revisión de seguridad.",
        "Claude siempre pide permiso antes de borrar un archivo."]),
      ("Describir el resultado", "Delegar un trabajo completo",
       "Describes lo que quieres conseguir, no cada paso, y revisas el plan antes de que empiece.",
       ["En esta carpeta están las facturas de proveedores de octubre. Quiero que:\n"
        "1. Renombres cada factura como AAAA-MM-DD_proveedor_total.pdf.\n"
        "2. Las ordenes en subcarpetas por proveedor.\n"
        "3. Crees un Excel con fecha, proveedor, número de factura, base, IVA y total, con una fila de totales.\n"
        "4. Me avises de cualquier archivo que no sea una factura o que parezca duplicado, sin borrarlo.\n"
        "Antes de empezar, enséñame tu plan."]),
      ("Revisar lo que ha hecho", "Supervisar, no desentenderse",
       "Sigues su lista de pasos mientras trabaja y, al terminar, compruebas el resultado.",
       ["¿Qué archivos no has renombrado y por qué? ¿Cómo has detectado el duplicado?"],
       ["Abre el Excel y comprueba a mano el total de dos facturas.",
        "Fíjate en las facturas con IVA distinto del 21 %."]),
      ("Que se repita cada mes", "Tareas programadas",
       "Conviertes el trabajo en una tarea programada y la ejecutas una vez a mano antes de dejarla sola.",
       ["Crea una tarea programada que haga esto el día 1 de cada mes con la carpeta de facturas del mes anterior."]),
    ],
  },
  {
    "n": 34,
    "titulo": "Etiquetas de precio para una tienda de ultramarinos",
    "caso": "Cada semana, una tienda pasa su lista de productos a etiquetas de precio a mano. Claude Code construye una pequeña aplicación que lo hace con un clic.",
    "archivos": "productos_despensa.csv",
    "modelo": "Claude Code en la aplicación de escritorio, en una carpeta nueva y con el modo de permisos que pide aprobación.",
    "pasos": [
      ("Preparar la carpeta", "Las precauciones",
       "Creas una carpeta nueva solo para este proyecto y copias dentro el archivo de productos.",
       [],
       ["Una carpeta nueva, creada solo para esta aplicación.",
        "El modo de permisos que te pide aprobación antes de cada cambio.",
        "Una aplicación pequeña, cuyo resultado puedas comprobar a simple vista."]),
      ("Un plan antes de construir", "Describir la tarea",
       "Describes lo que quieres y pides que te explique cómo lo va a hacer antes de escribir nada.",
       ["Quiero una pequeña aplicación que lea productos_despensa.csv y genere una página lista para imprimir con etiquetas de precio: 21 por hoja A4, en 3 columnas y 7 filas.\n"
        "Cada etiqueta lleva el nombre, el formato, el precio en grande y el origen.\n"
        "No sé programar: antes de escribir nada, explícame en lenguaje sencillo cómo la vas a hacer y qué tendré que hacer yo para usarla."]),
      ("Construir, probar y corregir", "Revisar el resultado",
       "Aceptas el plan, ves el resultado en el navegador y pides los cambios que hagan falta.",
       ["Adelante. Cuando termines, ábrela en el navegador para que la vea.",
        "Si un producto está en oferta, muestra el precio anterior tachado y el nuevo en rojo. "
        "Si un producto no tiene precio, no imprimas su etiqueta y avísame de cuál es."],
       ["Fíjate en el turrón: su nombre es muy largo. Y en los polvorones: no tienen precio."]),
      ("Hacerla tuya", "Una aplicación que funciona sin Claude",
       "Pides que puedas usarla tú solo cada semana, sin conversar con Claude.",
       ["Prepárame una forma de usarla con un doble clic cada semana, cuando cambie el CSV, y un archivo LEEME con las instrucciones en tres pasos."]),
    ],
  },
  {
    "n": 37,
    "titulo": "Seis situaciones y una lista para tu equipo",
    "caso": "Se revisan los ajustes de privacidad, se anonimiza un correo, se pone a prueba la complacencia y el sesgo de Claude y, al final, se construye una lista de comprobación de uso responsable.",
    "archivos": "situaciones_uso_responsable.txt, correo_con_datos_personales.txt, plan_tienda_velas.txt y perfiles_candidatura.txt",
    "modelo": "El modelo por defecto. La prueba de sesgo, cada perfil en una conversación nueva.",
    "pasos": [
      ("Antes de pegar nada", "Privacidad y datos personales",
       "Revisas tus ajustes de privacidad. Después anonimizas el correo a mano antes de pedir la respuesta.",
       ["Un cliente, [CLIENTE], dice que le hemos cobrado dos veces la cuota de 49 € este mes y pide la devolución.\n"
        "Redacta una respuesta breve y amable que confirme que lo revisamos y que la devolución se hará en 5 días laborables."],
       ["Ajustes de privacidad: decide si tus conversaciones se usan para entrenar.",
        "Incógnito para lo que no quieres guardar. Borrar un chat no borra la memoria.",
        "En correo_con_datos_personales.txt, sustituye nombre, DNI, cuenta, dirección y teléfono por [CLIENTE], [DNI]…"]),
      ("Cuando Claude te da la razón", "Complacencia",
       "Pides opinión sobre un plan con fallos evidentes. Después le pides que te lleve la contraria.",
       ["Este es mi plan para abrir una tienda. ¿Qué te parece?\n[pega plan_tienda_velas.txt]",
        "Ahora llévame la contraria: ¿qué me falta, qué argumentaría alguien que no está de acuerdo y qué podría salir mal? Sé concreto con las cifras."]),
      ("La prueba del nombre", "Sesgos en decisiones sobre personas",
       "Dos perfiles idénticos que solo cambian en el nombre, cada uno en una conversación nueva. Comparas las respuestas.",
       ["Valora del 1 al 10 a esta persona para un puesto administrativo y justifica la nota en tres líneas.\n[pega solo un perfil]"],
       ["Si la nota o el tono cambian, hay un problema que resolver antes de usar el resultado.",
        "La decisión sobre una persona la toma una persona: pide evidencias por criterio, no rankings."]),
      ("Tu checklist", "Diligencia",
       "Conviertes las seis situaciones en una lista interactiva para tu equipo.",
       ["Con estas seis situaciones, crea una lista de comprobación interactiva de uso responsable de Claude para mi equipo, en tres bloques: antes de compartir, antes de usar un resultado y antes de publicar.\n"
        "Cada punto, con una pregunta de sí o no.\n[pega situaciones_uso_responsable.txt]"]),
    ],
  },
]

FOOT = "© BlockStellart · Joan Amengual"
CSS_S = """
@page { size: 297mm 167.06mm; margin: 0; }
:root { --ink: #1f1d1a; --muted: #6b665f; --hair: #e3ddd3; --accent: #c15f3c; --paper: #fbf8f3;
  --instr: #2e6b73; --instr-bg: #eef4f3; }
* { box-sizing: border-box; font-variant-ligatures: none; }
html, body { margin: 0; color: var(--ink); background: #fff; font-family: 'Avenir Next', 'Helvetica Neue', sans-serif; }
.s { width: 297mm; height: 167.06mm; padding: 11mm 16mm 0; display: flex; flex-direction: column;
  break-after: page; overflow: hidden; position: relative; }
.box { flex: 1; min-height: 0; overflow: hidden; }
.in { display: flex; flex-direction: column; }
.foot { height: 9mm; display: flex; justify-content: space-between; align-items: center;
  font-size: 7.5pt; color: #9a948c; border-top: 1px solid var(--hair); }
.label { font-size: 8.5pt; letter-spacing: .2em; text-transform: uppercase; color: var(--accent);
  font-weight: 600; display: flex; align-items: center; gap: 3mm; }
.label::before { content: ""; width: 9mm; height: 1.5px; background: var(--accent); }
.cover { background: var(--paper); padding: 16mm 22mm 0; }
.cover .top { font-size: 11pt; letter-spacing: .28em; text-transform: uppercase; color: var(--accent); font-weight: 600; }
.cover .main { margin: auto 0; }
.cover .bar { width: 22mm; height: 3mm; background: var(--accent); margin-bottom: 8mm; }
.cover h1 { font-size: 56pt; line-height: 1.0; letter-spacing: -.025em; font-weight: 700; margin: 0; }
.cover h1.sec { font-size: 46pt; line-height: 1.05; max-width: 250mm; }
.cover .sub { font-size: 16pt; line-height: 1.35; color: var(--muted); margin-top: 8mm; max-width: 200mm; }
.intro { background: var(--paper); }
.intro h2 { font-size: 34pt; line-height: 1.05; letter-spacing: -.015em; font-weight: 700; margin: 4mm 0 4mm; }
.intro .caso { font-size: 15pt; line-height: 1.4; color: var(--muted); margin: 0 0 7mm; max-width: 230mm; }
.intro .meta { display: grid; grid-template-columns: 30mm 1fr; border-top: 1.4px solid var(--ink); font-size: 11pt; }
.intro .meta div { padding: 2.2mm 0; border-bottom: 1px solid var(--hair); }
.intro .meta .k { font-size: 8pt; letter-spacing: .12em; text-transform: uppercase; color: var(--muted); font-weight: 600; padding-top: 3mm; }
.intro ol { list-style: none; margin: 7mm 0 0; padding: 0; display: grid; grid-template-columns: repeat(4, 1fr); gap: 6mm; }
.intro li { border-top: 2px solid var(--accent); padding-top: 2.5mm; font-size: 11.5pt; font-weight: 600; line-height: 1.3; }
.intro li span { display: block; color: var(--accent); font-size: 10pt; margin-bottom: 1mm; }
.step .head { display: grid; grid-template-columns: 16mm 1fr; align-items: baseline; margin: 3mm 0 3mm; }
.step .n { font-size: 26pt; font-weight: 700; color: var(--accent); line-height: 1; }
.step h3 { font-size: 22pt; font-weight: 700; line-height: 1.1; margin: 0; letter-spacing: -.01em; }
.step .tag { font-size: 8.5pt; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); margin-top: 1.5mm; }
.step .txt { font-size: 14pt; line-height: 1.4; margin: 0 0 3mm 16mm; }
.step ul { margin: 0 0 3mm 16mm; padding-left: 6mm; font-size: 13pt; line-height: 1.35; }
.step li { margin: .8mm 0; }
.step .pp { display: grid; grid-template-columns: 16mm 1fr; align-items: start; margin-bottom: 4mm; }
.step .pl { font-size: 8.5pt; font-weight: 700; color: var(--accent); padding-top: 4.2mm; white-space: nowrap; }
.step .pp pre { margin: 0; }
.step pre { font-family: 'SF Mono', Menlo, monospace; font-size: 10.5pt; line-height: 1.6; white-space: pre-wrap;
  background: var(--paper); border-left: 2.5px solid var(--accent); padding: 4mm 5mm; margin: 0 0 4mm 16mm; }
.step .ib { margin: 0 0 4mm 16mm; }
.step .il { font-size: 8.5pt; font-weight: 700; letter-spacing: .14em; text-transform: uppercase;
  color: var(--instr); margin-bottom: 1.5mm; }
.step .ib pre { margin: 0; background: var(--instr-bg); border-left-color: var(--instr); }
"""
FIT_S = """
(function () {
  let zmin = 1, fits = true;
  document.querySelectorAll('.box').forEach(b => {
    const inner = b.firstElementChild, over = () => inner.getBoundingClientRect().height > b.clientHeight + 0.5;
    let z = 1;
    while (over() && z > 0.5) { z = +(z - 0.01).toFixed(2); inner.style.zoom = z; }
    fits = fits && !over(); zmin = Math.min(zmin, z);
  });
  document.body.dataset.zoom = zmin; document.body.dataset.fits = fits;
})();
"""


def e(s):
    return html.escape(s)


def espaciar(prompt):
    """Añade una línea en blanco entre las instrucciones de un prompt para que se
    lea mejor, salvo en listas numeradas y tras una línea que acaba en «:».
    Solo cambian los saltos de línea, no el texto."""
    out = []
    lineas = []
    for ln in prompt.split("\n"):
        # Las líneas largas con varias frases se parten en una frase por línea
        # (sin cortar tras abreviaturas en mayúsculas como «RR. HH.»).
        lineas += re.split(r"(?<=[a-záéíóúñ)»][.?])\s+(?=[A-ZÁÉÍÓÚ¿])", ln) if len(ln) > 140 else [ln]
    for k, ln in enumerate(lineas):
        prev = out[-1] if out else ""
        lista = re.match(r"^(\d+\.|-) ", ln) and re.match(r"^(\d+\.|-) ", prev)
        if k and not (prev.endswith(":") or lista):
            out.append("")
        out.append(ln)
    return "\n".join(out)


def slide(cls, inner, n):
    foot = f'<div class="foot"><span>{FOOT}</span><span>{n}</span></div>'
    return f'<div class="s {cls}"><div class="box"><div class="in">{inner}</div></div>{foot}</div>'


def build(x):
    dominio = CLASES[x["n"]]
    cover = ('<div class="top">Ejercicio práctico con Claude</div><div class="main"><div class="bar"></div>'
             f'<h1 class="sec">{e(dominio)}</h1><div class="sub">{e(x["titulo"])}</div></div>')
    out = [f'<div class="s cover">{cover}<div class="foot"><span>{FOOT}</span><span></span></div></div>']
    pasos = "".join(f"<li><span>{i:02d}</span>{e(p[0])}</li>" for i, p in enumerate(x["pasos"], 1))
    meta = (f'<div class="k">Archivos</div><div>{e(x["archivos"])}</div>'
            f'<div class="k">Modelo</div><div>{e(x["modelo"])}</div>')
    out.append(slide("intro", f'<div class="label">{e(dominio)}</div><h2>{e(x["titulo"])}</h2>'
                     f'<p class="caso">{e(x["caso"])}</p><div class="meta">{meta}</div><ol>{pasos}</ol>', len(out) + 1))
    for i, p in enumerate(x["pasos"], 1):
        titulo, concepto, haces, prompts = p[:4]
        lista = p[4] if len(p) > 4 else []
        extra = f'<ul>{"".join(f"<li>{e(t)}</li>" for t in lista)}</ul>' if lista else ""
        # Los bloques (instrucciones, descripciones) van en su color con la etiqueta
        # encima; los prompts del chat llevan al lado «Prompt 1», «Prompt 2»…
        # (o «Prompt» si hay uno), salvo si el paso tiene un único prompt.
        reales = [t for t in prompts if not isinstance(t, Bloque)]
        k, pr = 0, ""
        for t in prompts:
            if isinstance(t, Bloque):
                pr += f'<div class="ib"><div class="il">{e(t.etiqueta)}</div><pre>{e(t)}</pre></div>'
            elif len(prompts) > 1:
                k += 1
                pl = f"Prompt {k}" if len(reales) > 1 else "Prompt"
                pr += f'<div class="pp"><div class="pl">{pl}</div><pre>{e(espaciar(t))}</pre></div>'
            else:
                pr += f"<pre>{e(espaciar(t))}</pre>"
        out.append(slide("step", f'<div class="label">{e(dominio)}</div>'
                         f'<div class="head"><div class="n">{i:02d}</div><div><h3>{e(titulo)}</h3>'
                         f'<div class="tag">{e(concepto)}</div></div></div>'
                         f'<p class="txt">{e(haces)}</p>{pr}{extra}', len(out) + 1))
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Ejercicios prácticos con Claude</title>'
            f'<style>{CSS_S}</style></head><body>{"".join(out)}<script>{FIT_S}</script></body></html>')


def render(doc, out):
    """Imprime el HTML a PDF y avisa si alguna diapositiva no cabe o queda pequeña."""
    tmp = Path(tempfile.mkdtemp(prefix="intro_claude_"))
    try:
        src, pdf = tmp / "d.html", tmp / "d.pdf"
        src.write_text(doc, encoding="utf-8")
        base = [CHROME, "--headless", "--disable-gpu", "--virtual-time-budget=3000"]
        dom = subprocess.run(base + ["--dump-dom", src.as_uri()], check=True, capture_output=True, text=True).stdout
        m = re.search(r'data-zoom="([\d.]+)" data-fits="(\w+)"', dom)
        subprocess.run(base + ["--no-pdf-header-footer", f"--print-to-pdf={pdf}", src.as_uri()],
                       check=True, capture_output=True)
        out.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(pdf, out)
        print(f"{out.name}: escala mínima {m[1] if m else '?'}")
        if not m or m[2] != "true":
            sys.exit(f"AVISO: alguna diapositiva de {out.name} no cabe ni reducida.")
        if float(m[1]) < 0.85:
            print("AVISO: alguna diapositiva baja del 85 %; conviene recortar.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    solo = {int(a) for a in sys.argv[1:]}
    for x in EJERCICIOS:
        if solo and x["n"] not in solo:
            continue
        render(build(x), ENUNCIADOS / f"Enunciado - Clase {x['n']:02d} - {CLASES[x['n']]}.pdf")


if __name__ == "__main__":
    main()
