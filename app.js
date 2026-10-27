// ============================================================
//  app.js — UI y estado · Indicadores
//  Las reglas viven en logic.js (funciones puras) y el Excel en xlsx.js.
//  Aquí: datos en localStorage, enrutado por hash, asistente de 6 pasos,
//  cuadro de indicadores, compuesto, ficha/informe, aprende, glosario y quiz.
// ============================================================
(function () {
  "use strict";
  const L = window.Logica;

  // ---------- Versión y novedades ----------
  // Al publicar cambios: añadir una entrada AL PRINCIPIO de NOVEDADES, poner su versión en los ?v= de index.html, en
  // APP_VERSION y subir CACHE_NAME en sw.js. Formato de versión: AAAA.NN.
  const APP_VERSION = "2026.14";
  const NOVEDADES = [
    // [eco]
    { version: "2026.14", fecha: "Octubre 2026", titulo: "Nueva app hermana: ACR · Ishikawa", cambios: [
      "Enlace a la nueva app hermana ACR · Ishikawa en la portada y en «Acerca de»: busca las causas de fondo de un problema o de un incidente con el diagrama de Ishikawa (espina de pescado) y el análisis causa raíz, sin buscar culpables; sus acciones de mejora llevan indicador."] },
    { version: "2026.13", fecha: "Octubre 2026", titulo: "Traer de Lectura Crítica", cambios: [
      "En el paso 1, «Traer de Lectura Crítica» escribe el criterio a partir de la pregunta PICO de un artículo que hayas leído con esa app (si aún no tienes criterio) y guarda su evidencia en «Observaciones». Avisa si la decisión fue esperar o descartar."] },
    { version: "2026.12", fecha: "Octubre 2026", titulo: "Nueva app hermana: Lectura Crítica", cambios: [
      "Enlace a la nueva app hermana Lectura Crítica en la portada y en «Acerca de»: lee críticamente un artículo paso a paso (plantillas de CASPe con preguntas propias, calculadoras y un veredicto con semáforo) antes de cambiar la práctica."] },
    { version: "2026.11", fecha: "Octubre 2026", titulo: "Nueva herramienta de apoyo: Legibilidad", cambios: [
      "En «Acerca de», el grupo «Herramientas de apoyo» suma Legibilidad: mide con el índice de Flesch-Szigriszt y la escala INFLESZ si un texto para pacientes se entiende y te dice qué cambiar."] },
    { version: "2026.10", fecha: "Octubre 2026", titulo: "La copia sin conexión de las otras apps ya no se borra", cambios: [
      "Arreglado un fallo: al actualizarse esta app se borraba la copia sin conexión de las demás apps del ecosistema (y al revés). Ahora cada app solo limpia la suya, así que todas siguen funcionando sin conexión."] },
    { version: "2026.09", fecha: "Octubre 2026", titulo: "Herramientas de apoyo y portal de las apps", cambios: [
      "En «Acerca de», nuevo grupo «Herramientas de apoyo»: el Anonimizador RGPD (seudonimiza o anonimiza una hoja de cálculo sin que los datos salgan de tu navegador) y el Normalizador Bibliográfico (pon tus referencias en Vancouver, APA y otros estilos).",
      "Y un enlace al portal cjgaland.github.io, la página que reúne todas las apps y explica cómo encajan."] },
    { version: "2026.08", fecha: "Octubre 2026", titulo: "Nueva app hermana: Rondas de Seguridad", cambios: [
      "Acceso directo desde la portada y desde «Acerca de» a Rondas de Seguridad, la nueva app hermana: programa y haz las rondas de seguridad de una unidad, con su cumplimiento ronda a ronda y las acciones con su indicador."] },
    { version: "2026.07", fecha: "Octubre 2026", titulo: "Nueva app hermana: Mapa de Riesgos", cambios: [
      "Acceso directo desde la portada y desde «Acerca de» a Mapa de Riesgos, la nueva app hermana: elabora paso a paso el mapa de riesgos de un proceso (flujograma, AMFE, acciones de mejora y el mapa coloreado por IPR) y puede usar tus indicadores para medir las acciones."] },
    { version: "2026.06", fecha: "Octubre 2026", titulo: "Nuevas apps hermanas: Flujograma y DAFO · CAME", cambios: [
      "Acceso directo desde la portada y desde «Acerca de» a Flujograma (dibuja el flujograma de un proceso y marca sus puntos críticos) y a DAFO · CAME (análisis DAFO de una unidad y plan de acción con indicadores), las nuevas apps hermanas."] },
    { version: "2026.05", fecha: "Octubre 2026", titulo: "Nueva app hermana: AMFE", cambios: [
      "Acceso directo desde la portada y desde «Acerca de» a AMFE (Análisis Modal de Fallos y Efectos), la app hermana para analizar los fallos de un proceso, puntuar G·O·D e IPR y priorizar. Con su botón «Traer de Indicadores» usa los indicadores que tengas aquí para medir sus acciones de mejora."] },
    { version: "2026.04", fecha: "Octubre 2026", titulo: "Nueva app hermana: Auditorías", cambios: [
      "Acceso directo desde la portada y desde «Acerca de» a Auditorías, la app hermana para planificar, realizar, informar y hacer el seguimiento de una auditoría completa: Mapa de Riesgos, Plan Anual, diseño con indicadores y muestra, recogida de datos, informe y plan de mejora."] },
    // [/eco]
    /* [ind]
    { version: "2026.13", fecha: "Octubre 2026", titulo: "Evidencia de Lectura Crítica", cambios: [
      "En el paso 1, «Cargar archivo de Lectura Crítica» escribe el criterio a partir de la pregunta PICO de un artículo leído con la app Lectura Crítica (desde su copia de seguridad .json), si aún no tienes criterio, y guarda su evidencia en «Observaciones». Avisa si la decisión fue esperar o descartar."] },
    [/ind] */
    { version: "2026.03", fecha: "Octubre 2026", titulo: "Una app para todos", cambios: [
      "La app es ahora de uso general: se han quitado las referencias directas al curso de auditorías (que sigue citado entre las fuentes, en «Acerca de»).",
      // [eco]
      "La portada se titula «Indicadores» y tiene un acceso directo a la app hermana Tamaño Muestral.",
      // [/eco]
      /* [ind]
      "La portada se titula «Indicadores» y tiene un acceso directo a la calculadora Tamaño Muestral.",
      [/ind] */
      "El Excel exportado nombra sus hojas «Caracterización», «Definición operativa», «Compuestos» y «Resultados».",
      "Nuevo «Cuadro de ejemplo» (menú «Ejemplos»), aparte de tus indicadores: 9 indicadores y un compuesto con resultados inventados para ver semáforos, gráficos y tendencia. No se guarda ni se mezcla con lo tuyo; «Copiar a mi cuadro» trae el que te sirva.",
      "Si tu cuadro tenía copias de los ejemplos de versiones anteriores, un aviso te permite quitarlas.",
      "Gráfico de resultados: las fechas ya no se montan cuando hay muchos periodos."] },
    { version: "2026.02", fecha: "Octubre 2026", titulo: "Resultados por periodo (Verificar)", cambios: [
      "Nueva pantalla «Resultados» en cada indicador: escribe un periodo por fila y verás si cumple el estándar (verde o rojo) según el sentido (≥ o ≤).",
      "Gráfico de la evolución con la línea del estándar y un resumen de la serie: en cuántos periodos se cumple, rachas en rojo y si mejora o empeora.",
      "En los indicadores LQAS se aplica la regla n/c del plan y se enlaza con «Ya tengo los resultados» de Tamaño Muestral para la frase del informe.",
      "Los resultados salen en la ficha imprimible, en el cuadro, en el Excel (hoja «Resultados») y en la copia de seguridad. Los ejemplos traen mediciones inventadas."] },
    { version: "2026.01", fecha: "Octubre 2026", titulo: "Primera versión", cambios: [
      "Asistente en seis pasos para crear un indicador: del hallazgo del Mapa de Riesgos al indicador, caracterización, definición operativa, comprobación de su calidad (SMART, válido, fiable, útil), muestra y ficha.",
      "Avisos de las trampas más habituales: centinela sin muestreo, la media que maquilla, el sentido invertido del estándar, términos sin definir…",
      "Indicador compuesto con su desglose, el peso de cada criterio y la frase «titular + desglose».",
      "Cuadro de indicadores imprimible y exportable a Excel; copia de seguridad e importación; enlace para compartir un indicador.",
      "Enlace a Tamaño Muestral con los datos ya puestos para calcular la muestra, el plan LQAS o la concordancia del pilotaje.",
      "Aprende, glosario con buscador y «Ponte a prueba»."] }
  ];

  // ---------- Utilidades ----------
  const nf = (x, d = 1) => Number.isFinite(x) ? x.toLocaleString("es-ES", { maximumFractionDigits: d, minimumFractionDigits: 0 }) : "—";
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const escMl = s => esc(s).replace(/\n/g, "<br>");
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const view = $("#view");
  const hoy = () => new Date().toISOString();
  const fechaCorta = iso => { const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleDateString("es-ES"); };
  const genId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const lleno = L.lleno;
  const clon = o => JSON.parse(JSON.stringify(o));
  const leer = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const escribir = (k, v) => { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } };

  // ---------- Datos (solo en este dispositivo) ----------
  const KEY_DATOS = "indicadores-datos";
  let datos = { indicadores: [], compuestos: [], ultimaCopia: null };
  (function cargar() {
    const crudo = leer(KEY_DATOS); if (!crudo) return;
    try {
      const o = JSON.parse(crudo), s = L.sanearDatos(o, genId, hoy());
      if (s) datos = { indicadores: s.indicadores, compuestos: s.compuestos, ultimaCopia: typeof o.ultimaCopia === "string" ? o.ultimaCopia : null };
    } catch (e) { /* datos dañados: se empieza de cero sin borrar lo guardado */ }
  })();
  // Los datos del usuario son los únicos que se guardan. Los ejemplos viven aparte, en memoria
  // («Cuadro de ejemplo», ids «ej-…»): se pueden tocar, pero ni se guardan ni se mezclan con los suyos.
  const datosUsuario = datos;
  let datosEjemplo = null, modoEjemplo = false;
  const esIdEjemplo = id => /^ej-/.test(id || "");
  function cuadroEjemplo() {
    if (!datosEjemplo) datosEjemplo = { indicadores: EJEMPLOS.map((e, i) => indDesdeEjemplo(e, "ej-" + (i + 1))), compuestos: [Object.assign(L.nuevoCompuesto("ej-c1", hoy()), clon(EJEMPLO_COMPUESTO))], ultimaCopia: null };
    return datosEjemplo;
  }
  const rutaCuadro = () => modoEjemplo ? "#/ejemplos" : "#/cuadro";
  const volverCuadro = () => `<a class="volver-cuadro" href="${rutaCuadro()}">← ${modoEjemplo ? "Cuadro de ejemplo" : "Mi cuadro"}</a>`;
  let tGuardar = null, avisoSinEspacio = false;
  function guardar(inmediato) {
    clearTimeout(tGuardar);
    const fn = () => {
      const ok = escribir(KEY_DATOS, JSON.stringify({ app: "indicadores", v: L.VERSION_DATOS, ultimaCopia: datosUsuario.ultimaCopia, indicadores: datosUsuario.indicadores, compuestos: datosUsuario.compuestos }));
      if (!ok && !avisoSinEspacio) { avisoSinEspacio = true; toast("No se ha podido guardar en este navegador (¿modo privado o sin espacio?). Haz una copia de seguridad.", 6000); }
    };
    if (inmediato) fn(); else tGuardar = setTimeout(fn, 250);
  }
  window.addEventListener("pagehide", () => guardar(true));
  const buscaInd = id => datos.indicadores.find(x => x.id === id);
  const buscaCmp = id => datos.compuestos.find(x => x.id === id);
  // Un borrador no se guarda hasta que se escribe algo
  function asegurarGuardado(obj, lista) { if (!lista.includes(obj)) lista.push(obj); }

  // ---------- Textos de ayuda reutilizables ----------
  const box = (tipo, tag, html) => `<div class="box box--${tipo}"><span class="box-tag">${tag}</span>${html}</div>`;
  const idea = html => box("idea", "Idea clave", `<p>${html}</p>`);
  const trampa = html => box("trampa", "Ojo, trampa", `<p>${html}</p>`);
  const practica = html => box("practica", "En la práctica", `<p>${html}</p>`);
  const ejemplo = (titulo, html) => box("ejemplo", "Ejemplo resuelto" + (titulo ? " · " + titulo : ""), html);
  const KEY_FACIL = "indicadores-facil-abierto";
  const facil = html => `<details class="facil"${leer(KEY_FACIL) === "1" ? " open" : ""}><summary>Explicado fácil</summary><div class="facil-body">${html}</div></details>`;

  // ---------- Campos ----------
  // path: "nombre", "muestra.e", "calidad.E.t"…
  function getPath(o, path) { return path.split(".").reduce((a, k) => (a == null ? undefined : a[k]), o); }
  function setPath(o, path, v) {
    const ks = path.split("."), ult = ks.pop();
    let a = o; ks.forEach(k => { if (a[k] == null || typeof a[k] !== "object") a[k] = {}; a = a[k]; });
    a[ult] = v;
  }
  let nAyuda = 0;
  function ayudaBtn(help) {
    if (!help) return { btn: "", box: "" };
    const id = "h" + (++nAyuda);
    return { btn: `<button type="button" class="info-btn" data-info="${id}" aria-expanded="false" aria-label="Ayuda">i</button>`, box: `<p class="help" id="help-${id}" hidden>${help}</p>` };
  }
  // Campo de texto (una línea o varias)
  function campo(obj, path, label, o = {}) {
    const v = getPath(obj, path) || "", a = ayudaBtn(o.help);
    const ctrl = o.area
      ? `<textarea class="txt-in" data-k="${path}" rows="${o.rows || 2}" placeholder="${esc(o.ph || "")}">${esc(v)}</textarea>`
      : `<div class="input-wrap${o.num ? " input-wrap--num" : ""}"><input class="txt-in" data-k="${path}" value="${esc(v)}" placeholder="${esc(o.ph || "")}"${o.num ? ' inputmode="decimal"' : ""} autocomplete="off" />${o.unit ? `<span class="input-unit">${esc(o.unit)}</span>` : ""}</div>`;
    const quick = o.quick ? `<div class="quick">${o.quick.map(q => `<button type="button" class="chip${String(v) === String(q) ? " active" : ""}" data-set="${path}" data-v="${esc(q)}">${esc(q)}${o.unit === "%" ? " %" : ""}</button>`).join("")}</div>` : "";
    return `<div class="field${o.destacado ? " field--destacado" : ""}"><div class="field-head"><label class="field-label">${label}</label>${o.opt ? '<span class="field-opt">(opcional)</span>' : ""}${a.btn}</div>${a.box}${o.sub ? `<p class="field-sub">${o.sub}</p>` : ""}${ctrl}${quick}${o.extra || ""}</div>`;
  }
  // Grupo de botones excluyentes
  function seg(obj, path, opciones) {
    const v = getPath(obj, path) || "";
    return `<div class="seg" role="group">${opciones.map(([val, txt]) => `<button type="button" class="${v === val ? "active" : ""}" data-pick="${path}" data-v="${val}">${txt}</button>`).join("")}</div>`;
  }

  // ---------- Ejemplos (de la parte teórica del curso o inventados; NO de los ejercicios prácticos) ----------
  // Mediciones inventadas para enseñar la fase «Verificar»
  const R = (periodo, num, den, nota = "") => ({ periodo, num, den, valor: "", nota });
  const cal = (t) => Object.fromEntries(L.CALIDAD.map(q => [q.k, { v: "si", t: t[q.k] || "" }]));
  const EJEMPLOS = [
    { titulo: "Valoración del riesgo de UPP", sub: "Proceso · porcentaje · LQAS", ind: {
      subproceso: "Valoración de enfermería al ingreso", hallazgo: "No se valora el riesgo de úlceras por presión al ingreso", g: "7", o: "6", d: "5",
      criterio: "Se registra la valoración del riesgo de úlceras por presión (escala Braden) en las primeras 24 horas del ingreso",
      nombre: "% de pacientes ingresados con la valoración del riesgo de UPP registrada en las primeras 24 horas",
      tipo: "proceso", justTipo: "Comprueba si se hace una actividad que marca el protocolo (valorar el riesgo), no el resultado final (que aparezcan o no úlceras).", calculo: "porcentaje",
      numerador: "pacientes con la escala Braden completa registrada en las primeras 24 h", denominador: "pacientes ingresados más de 24 h revisados",
      poblacion: "Pacientes ingresados en la unidad de hospitalización en el trimestre", exclusiones: "Estancias de menos de 24 h. Pacientes trasladados desde otra unidad con la valoración ya registrada.",
      fuentes: ["Historia clínica"], sentido: "mas", estandar: "95", umbral: "75", periodicidad: "Trimestral", responsable: "Supervisión de enfermería de la unidad",
      terminos: "«Registrada»: escala Braden con sus seis ítems y la puntuación total en el formulario de valoración. «24 horas»: desde la hora de llegada a la unidad que consta en admisión.",
      observaciones: "Se presenta en la Comisión de Calidad cada trimestre. Si se rechaza el lote, se revisan las historias que no cumplen para buscar causas.",
      calidad: cal({ E: "Un solo criterio, en una población y un plazo concretos.", M: "Cumple / no cumple en cada historia.", A: "Se revisan unas pocas historias al trimestre.", R: "Sale de un riesgo priorizado del Mapa de Riesgos.", T: "Trimestral.", valido: "Mide la valoración hecha, no solo que exista el formulario.", fiable: "Términos definidos y pilotado con dos auditoras.", util: "La supervisión puede reforzar la formación y el recordatorio.", comparable: "Mismo método cada trimestre y entre unidades." }),
      piloto: { a: "24", b: "1", c: "1", d: "4" },
      objetivo: "decidir", muestra: { conf: "95", e: "10", p: "", N: "", loss: "", alpha: "5", beta: "20", sigma: "", eMedia: "" },
      plan: { n: "16", c: "14", nota: "Plan exacto de Tamaño Muestral (estándar 95 %, umbral 75 %, α 5 %, β 20 %)." },
    } },
    { titulo: "Conciliación de la medicación al ingreso", sub: "Proceso · porcentaje · estimar", ind: {
      subproceso: "Ingreso en hospitalización", hallazgo: "No se concilia la medicación habitual de los pacientes polimedicados al ingreso", g: "8", o: "6", d: "6",
      criterio: "Se realiza y documenta la conciliación de la medicación en las primeras 24 horas tras el ingreso",
      nombre: "% de historias con la conciliación de la medicación realizada y documentada en las primeras 24 horas tras el ingreso",
      tipo: "proceso", justTipo: "Valora si se hace una actividad clave del protocolo de seguridad del medicamento.", calculo: "porcentaje",
      numerador: "historias con la conciliación documentada en las primeras 24 h", denominador: "historias de pacientes polimedicados ingresados revisadas",
      poblacion: "Pacientes con 5 o más fármacos crónicos ingresados en el semestre", exclusiones: "Estancias de menos de 24 h. Pacientes que no pueden informar y sin cuidador ni registro previo.",
      fuentes: ["Historia clínica", "Sistema de información"], fuenteOtra: "Módulo de prescripción", sentido: "mas", estandar: "90", umbral: "", periodicidad: "Semestral", responsable: "Jefatura de servicio y farmacia hospitalaria",
      terminos: "«Conciliación documentada»: consta la lista de medicación habitual y, para cada fármaco, si se mantiene, se suspende o se cambia. «Polimedicado»: 5 o más fármacos crónicos.",
      observaciones: "Se acompaña de un indicador de resultado (eventos adversos por omisiones o duplicidades).",
      calidad: cal({}), objetivo: "estimar", muestra: { conf: "95", e: "10", p: "", N: "400", loss: "", alpha: "5", beta: "20", sigma: "", eMedia: "" },
      plan: { n: "78", c: "", nota: "Estimar con ±10 % y 95 % de confianza en un universo de 400 historias." },
    } },
    { titulo: "Eventos adversos por omisiones o duplicidades", sub: "Resultado · número absoluto", ind: {
      subproceso: "Ingreso en hospitalización", hallazgo: "No se concilia la medicación habitual de los pacientes polimedicados al ingreso", g: "8", o: "6", d: "6",
      criterio: "No se producen eventos adversos por omisiones o duplicidades de medicación",
      nombre: "Nº de eventos adversos relacionados con omisiones o duplicidades de medicación identificados al alta",
      tipo: "resultado", justTipo: "Mide lo que le pasa al paciente: el efecto final de una mala conciliación.", calculo: "absoluto",
      numerador: "eventos adversos por omisión o duplicidad de medicación identificados al alta", denominador: "",
      poblacion: "Altas de la unidad en el mes", exclusiones: "Discrepancias justificadas y documentadas por el médico.",
      fuentes: ["Notificación de incidentes", "Historia clínica"], sentido: "menos", estandar: "0", umbral: "", periodicidad: "Mensual", responsable: "Referente de seguridad del paciente de la unidad",
      terminos: "«Evento adverso»: daño al paciente (o necesidad de tratamiento o vigilancia extra) por la omisión o duplicidad de un fármaco.",
      observaciones: "Es un número absoluto: no se puede comparar con otras unidades de distinto tamaño. Si interesa compararlo, conviene pasarlo a tasa por cada 100 altas.",
      calidad: Object.assign(cal({ T: "Mensual.", util: "La unidad puede reforzar la conciliación." }), { comparable: { v: "dudoso", t: "Sin denominador: no sirve para comparar unidades de distinto tamaño." } }), objetivo: "sin", plan: { n: "", c: "", nota: "Se cuentan todos los eventos del mes." } } },
    { titulo: "Radiografías de tobillo (Ottawa)", sub: "Proceso · porcentaje · LQAS", ind: {
      subproceso: "Atención en urgencias", hallazgo: "Radiografías de tobillo pedidas sin indicación", g: "", o: "", d: "",
      criterio: "Las radiografías de tobillo por esguince se piden solo si se cumplen los criterios de Ottawa",
      nombre: "% de radiografías de tobillo pedidas por esguince que cumplen los criterios de Ottawa",
      tipo: "proceso", justTipo: "Valora la adecuación de una decisión clínica a una guía.", calculo: "porcentaje",
      numerador: "peticiones de radiografía que cumplen los criterios de Ottawa", denominador: "peticiones de radiografía de tobillo por esguince revisadas",
      poblacion: "Esguinces de tobillo atendidos en los últimos seis meses", exclusiones: "Menores de 18 años. Politraumatizados.",
      fuentes: ["Historia clínica"], sentido: "mas", estandar: "90", umbral: "75", periodicidad: "Semestral", responsable: "Coordinación médica de urgencias",
      terminos: "«Cumple Ottawa»: consta en la historia al menos uno de los criterios de Ottawa (dolor en maléolo con hipersensibilidad ósea o incapacidad para apoyar).",
      observaciones: "Las tablas clásicas de Lemeshow dan 33/26, pero su β real es mayor que el 20 %: el plan exacto es 40/33.",
      calidad: cal({}), objetivo: "decidir", muestra: { conf: "95", e: "10", p: "", N: "", loss: "", alpha: "5", beta: "20", sigma: "", eMedia: "" },
      plan: { n: "40", c: "33", nota: "Plan exacto de Tamaño Muestral (90 % / 75 %, α 5 %, β 20 %)." },
    } },
    { titulo: "Registro del dolor", sub: "Proceso · porcentaje · LQAS", ind: {
      subproceso: "Atención en urgencias", hallazgo: "No se registra la intensidad del dolor", g: "", o: "", d: "",
      criterio: "Se registra la intensidad del dolor (escala numérica de 0 a 10) en la primera valoración",
      nombre: "% de pacientes con la intensidad del dolor registrada en la primera valoración",
      tipo: "proceso", justTipo: "Comprueba que se hace un paso del proceso asistencial.", calculo: "porcentaje",
      numerador: "pacientes con la escala de dolor registrada en la primera valoración", denominador: "pacientes atendidos revisados",
      poblacion: "Pacientes adultos atendidos en el último trimestre", exclusiones: "Pacientes inconscientes o que no pueden comunicarse (se valoran con otra escala).",
      fuentes: ["Historia clínica"], sentido: "mas", estandar: "85", umbral: "60", periodicidad: "Trimestral", responsable: "Supervisión de enfermería de urgencias",
      terminos: "«Registrada»: consta un número de 0 a 10. «No dolor» escrito sin número no cuenta.",
      observaciones: "La tabla clásica de Lemeshow da 16/11; el plan exacto es 22/16.",
      calidad: cal({}), objetivo: "decidir", muestra: { conf: "95", e: "10", p: "", N: "", loss: "", alpha: "5", beta: "20", sigma: "", eMedia: "" },
      plan: { n: "22", c: "16", nota: "Plan exacto de Tamaño Muestral (85 % / 60 %, α 5 %, β 20 %)." } } },
    { titulo: "Cirugía en lugar equivocado", sub: "Resultado · centinela", ind: {
      subproceso: "Bloque quirúrgico", hallazgo: "Riesgo de intervenir en el lado o lugar equivocado", g: "10", o: "2", d: "4",
      criterio: "Ningún paciente es intervenido en un lugar equivocado",
      nombre: "Cirugía en lugar equivocado",
      tipo: "resultado", justTipo: "Es un daño grave al paciente: el resultado que se quiere evitar.", calculo: "centinela",
      numerador: "intervenciones realizadas en el lugar, lado o nivel equivocado", denominador: "",
      poblacion: "Todas las intervenciones quirúrgicas del centro", exclusiones: "",
      fuentes: ["Notificación de incidentes", "Historia clínica"], sentido: "menos", estandar: "0", umbral: "", periodicidad: "Continua (cada caso)", responsable: "Dirección médica y Unidad de Calidad",
      terminos: "Incluye lado (derecho/izquierdo), órgano, nivel vertebral o dedo equivocados, aunque se detecte durante la intervención.",
      observaciones: "Cada caso se notifica y se hace un análisis de causas (análisis causa raíz). No se muestrea.",
      calidad: cal({}), objetivo: "sin", plan: { n: "", c: "", nota: "Vigilancia de todos los casos." } } },
    { titulo: "Caídas por 1.000 estancias", sub: "Resultado · tasa", ind: {
      subproceso: "Cuidados en hospitalización", hallazgo: "Caídas de pacientes durante el ingreso", g: "7", o: "5", d: "5",
      criterio: "Se previenen las caídas de los pacientes hospitalizados",
      nombre: "Caídas con daño por cada 1.000 estancias",
      tipo: "resultado", justTipo: "Mide un daño al paciente.", calculo: "tasa",
      numerador: "caídas con daño notificadas", denominador: "estancias del periodo", base: "1.000",
      poblacion: "Pacientes ingresados en las unidades de hospitalización", exclusiones: "Pediatría y psiquiatría (se miden aparte).",
      fuentes: ["Notificación de incidentes", "Sistema de información"], sentido: "menos", estandar: "1", umbral: "", periodicidad: "Trimestral", responsable: "Comisión de Seguridad del Paciente",
      terminos: "«Caída con daño»: cualquier lesión, aunque sea leve, que conste en la notificación o en la historia. «Estancia»: cada día de ingreso (dato del sistema de información).",
      observaciones: "Estándar orientativo inventado para el ejemplo: compárate con tu propia serie y con unidades parecidas.",
      calidad: cal({}), objetivo: "sin", plan: { n: "", c: "", nota: "Se cuentan todas las caídas notificadas y todas las estancias." },
    } },
    { titulo: "Esperas de más de 30 minutos", sub: "Proceso · de media a porcentaje", ind: {
      subproceso: "Atención en urgencias", hallazgo: "Demora hasta la primera atención", g: "", o: "", d: "",
      criterio: "Los pacientes reciben la primera atención en menos de 30 minutos",
      nombre: "% de pacientes que esperan más de 30 minutos hasta la primera atención",
      tipo: "proceso", justTipo: "Mide cómo funciona el circuito de atención.", calculo: "porcentaje",
      numerador: "pacientes con más de 30 minutos entre la llegada y la primera atención", denominador: "pacientes atendidos en el periodo",
      poblacion: "Pacientes atendidos en urgencias", exclusiones: "Pacientes que se van sin ser atendidos (se miden aparte).",
      fuentes: ["Sistema de información"], sentido: "menos", estandar: "10", umbral: "", periodicidad: "Mensual", responsable: "Coordinación de urgencias",
      terminos: "«Llegada»: hora de registro en admisión. «Primera atención»: hora de la primera valoración que consta en el sistema.",
      observaciones: "Antes se usaba el tiempo medio de espera (20 minutos), que escondía pacientes de tres horas. Como porcentaje se ve cuántos quedan fuera del límite.",
      calidad: cal({}), objetivo: "sin", plan: { n: "", c: "", nota: "El sistema de información da todos los casos: no hace falta muestra." },
    } },
    { titulo: "Dispensadores de solución hidroalcohólica", sub: "Estructura · porcentaje", ind: {
      subproceso: "Higiene de manos", hallazgo: "Falta de solución hidroalcohólica en el punto de atención", g: "6", o: "5", d: "3",
      criterio: "Cada punto de atención dispone de solución hidroalcohólica con producto",
      nombre: "% de puntos de atención con dispensador de solución hidroalcohólica operativo",
      tipo: "estructura", justTipo: "Valora un recurso necesario para que se pueda hacer bien la higiene de manos.", calculo: "porcentaje",
      numerador: "puntos de atención con dispensador con producto y que funciona", denominador: "puntos de atención observados",
      poblacion: "Puntos de atención (cama, box o consulta) de la unidad", exclusiones: "",
      fuentes: ["Observación directa"], sentido: "mas", estandar: "100", umbral: "", periodicidad: "Mensual", responsable: "Supervisión de la unidad",
      terminos: "«Operativo»: tiene producto y dispensa al accionarlo. «Punto de atención»: lugar donde se atiende al paciente, a menos de un brazo.",
      observaciones: "Recorrido mensual con una hoja de observación.",
      calidad: cal({}), objetivo: "sin", plan: { n: "", c: "", nota: "Se recorren todos los puntos de atención." },
    } }
  ];
  const EJEMPLO_COMPUESTO = { nombre: "Atención al infarto de miocardio", pacientes: "15", estandar: "", conclusion: "", rows: [
    { nom: "Aspirina en las primeras 24 h", cum: "10", opo: "12", ind: "" }, { nom: "β-bloqueantes en las primeras 24 h", cum: "4", opo: "6", ind: "" },
    { nom: "Aspirina al alta", cum: "4", opo: "5", ind: "" }, { nom: "β-bloqueantes al alta", cum: "2", opo: "4", ind: "" },
    { nom: "Consejo antitabaco", cum: "1", opo: "3", ind: "" }, { nom: "Trombolíticos en los primeros 30 minutos", cum: "7", opo: "10", ind: "" }] };
  // Resultados inventados para los ejemplos (mismo orden que EJEMPLOS): verdes, rojos, rachas y tendencia
  const RESULTADOS_EJ = [
    [R("2025-T1", "15", "16"), R("2025-T2", "13", "16", "Faltaba la escala en los ingresos del fin de semana."), R("2025-T3", "12", "16", "Segundo trimestre seguido en rojo: se revisa el circuito de ingreso."), R("2025-T4", "15", "16", "Recordatorio en el cambio de turno y campo obligatorio en la historia."), R("2026-T1", "16", "16"), R("2026-T2", "14", "16"), R("2026-T3", "16", "16")],
    [R("2024-S1", "58", "78"), R("2024-S2", "63", "78", "Se forma al personal en conciliación."), R("2025-S1", "66", "78"), R("2025-S2", "71", "78", "Farmacia revisa los ingresos de polimedicados."), R("2026-S1", "72", "78", "Dato de una muestra: mira la horquilla en Tamaño Muestral antes de concluir.")],
    [R("Abr 2026", "2", ""), R("May 2026", "1", ""), R("Jun 2026", "1", ""), R("Jul 2026", "0", "", "Primer mes sin eventos tras implantar la conciliación al ingreso."), R("Ago 2026", "0", ""), R("Sep 2026", "1", "", "Un evento por duplicidad de un anticoagulante: se analiza con farmacia.")],
    [R("2024-S1", "31", "40", "Rojo: se pedían radiografías sin explorar los criterios."), R("2024-S2", "34", "40"), R("2025-S1", "36", "40"), R("2025-S2", "35", "40"), R("2026-S1", "31", "40", "Rojo: llegada de residentes nuevos sin formación en los criterios de Ottawa. Se programa una sesión.")],
    [R("2025-T1", "14", "22"), R("2025-T2", "15", "22", "Dos trimestres en rojo: se añade la escala de dolor al formulario de triaje."), R("2025-T3", "18", "22"), R("2025-T4", "19", "22"), R("2026-T1", "20", "22"), R("2026-T2", "17", "22"), R("2026-T3", "21", "22")],
    [R("2024-S1", "0", ""), R("2024-S2", "1", "", "Un caso (lado equivocado, detectado antes de la incisión): análisis causa raíz y refuerzo del marcado del sitio quirúrgico."), R("2025-S1", "0", ""), R("2025-S2", "0", ""), R("2026-S1", "0", "")],
    [R("2025-T1", "4", "2.900"), R("2025-T2", "3", "3.050"), R("2025-T3", "5", "2.980"), R("2025-T4", "4", "3.120", "Dos trimestres seguidos por encima del estándar."), R("2026-T1", "3", "2.950"), R("2026-T2", "2", "3.100", "Se añadió la valoración del riesgo de caídas al ingreso."), R("2026-T3", "1", "3.000")],
    [R("Abr 2026", "340", "2.410"), R("May 2026", "318", "2.450"), R("Jun 2026", "301", "2.520"), R("Jul 2026", "312", "2.480"), R("Ago 2026", "270", "2.390"), R("Sep 2026", "230", "2.510", "Segundo profesional en admisión en las horas punta.")],
    [R("Abr 2026", "44", "50"), R("May 2026", "45", "50"), R("Jun 2026", "47", "50", "Se nombra un responsable de reponer por turno."), R("Jul 2026", "46", "50"), R("Ago 2026", "49", "50"), R("Sep 2026", "50", "50")]
  ];
  function indDesdeEjemplo(e, id) {
    const ind = Object.assign(L.nuevoIndicador(id || genId(), hoy()), clon(e.ind));
    const i = EJEMPLOS.indexOf(e); if (RESULTADOS_EJ[i]) ind.resultados = clon(RESULTADOS_EJ[i]);
    ind.muestra = Object.assign(L.nuevoIndicador("x", "").muestra, e.ind.muestra || {});
    ind.piloto = Object.assign({ a: "", b: "", c: "", d: "" }, e.ind.piloto || {});
    return ind;
  }

  // ---------- Pasos del asistente ----------
  const PASOS = [
    { n: 1, corto: "Riesgo", titulo: "Del riesgo al indicador" },
    { n: 2, corto: "Tipo", titulo: "Caracterízalo" },
    { n: 3, corto: "Definición", titulo: "Definición operativa" },
    { n: 4, corto: "Calidad", titulo: "¿Es un buen indicador?" },
    { n: 5, corto: "Muestra", titulo: "¿Cuántos casos reviso?" },
    { n: 6, corto: "Ficha", titulo: "Ficha del indicador" }
  ];
  let actual = null, pasoActual = 1;   // indicador en edición

  function stepperHtml(ind, paso) {
    const comp = L.pasosCompletos(ind);
    return `<ol class="wiz-steps" aria-label="Pasos">${PASOS.map(p => `<li><a href="#/indicador/${ind.id}/${p.n}" class="wiz-step${p.n === paso ? " actual" : ""}${comp[p.n] ? " hecho" : ""}"${p.n === paso ? ' aria-current="step"' : ""}><span class="wiz-num">${comp[p.n] && p.n !== paso ? "✓" : p.n}</span><span class="wiz-lbl">${p.corto}</span></a></li>`).join("")}</ol>`;
  }

  function avisosHtml(ind, paso) {
    const A = L.avisos(ind).filter(a => a.paso === paso);
    return A.map(a => a.nivel === "nota" ? `<div class="note">ℹ ${esc(a.texto)}</div>` : a.nivel === "error" ? `<div class="aviso-error">✕ ${esc(a.texto)}</div>` : `<div class="warn">⚠ ${esc(a.texto)}</div>`).join("");
  }

  // Resumen en vivo (columna derecha en escritorio)
  function resumenHtml(ind) {
    const t = L.TIPOS[ind.tipo], c = L.CALCULOS[ind.calculo], ipr = L.ipr(ind.g, ind.o, ind.d), est = L.estadoCalidad(ind);
    const fila = (k, v) => v ? `<div class="res-fila"><dt>${k}</dt><dd>${v}</dd></div>` : "";
    return `<div class="resumen">
      <div class="res-label">Así va tu indicador</div>
      <h3 class="resumen-nombre">${lleno(ind.nombre) ? esc(ind.nombre) : '<span class="vacio">Sin nombre todavía</span>'}</h3>
      <div class="badges">${t ? `<span class="badge badge--${ind.tipo}">${t.nom}</span>` : ""}${c ? `<span class="badge">${c.nom}</span>` : ""}${ipr ? `<span class="badge badge--ipr">IPR ${ipr}</span>` : ""}<span class="badge badge--sem badge--${est.color}"><span class="punto"></span>${est.texto}</span></div>
      <dl class="res-dl">
        ${fila("Hallazgo", esc(ind.hallazgo))}
        ${fila("Fórmula", ind.calculo ? `<span class="formula-mini">${esc(L.formula(ind))}</span>` : "")}
        ${fila("Estándar", esc(L.textoEstandar(ind)))}
        ${fila("Periodicidad", esc(ind.periodicidad))}
        ${fila("Responsable", esc(ind.responsable))}
        ${fila("Plan", esc(L.textoPlan(ind)))}
      </dl>
    </div>`;
  }

  // ----- Paso 1 -----
  function paso1(ind) {
    const ipr = L.ipr(ind.g, ind.o, ind.d);
    return `
      ${idea("Un indicador nace de un <b>riesgo</b>: el Mapa de Riesgos dice qué puede fallar; el <b>criterio</b> dice cómo debería hacerse bien; el <b>indicador</b> lo convierte en algo que se puede contar.")}
      ${facil(`
        <h3>El camino en cuatro escalones</h3>
        <ol>
          <li><b>Subproceso</b>: la parte del proceso donde está el riesgo (ingreso, triaje, quirófano…).</li>
          <li><b>Hallazgo</b> (fallo del Mapa de Riesgos): lo que puede salir mal. En el AMFE cada fallo tiene su <b>IPR</b> = gravedad × ocurrencia × detección; se priorizan los IPR más altos.</li>
          <li><b>Criterio</b>: la regla de buena práctica que debería cumplirse, escrita en positivo y sin ambigüedad.</li>
          <li><b>Indicador</b>: el criterio convertido en un número que se puede contar y comparar con un objetivo.</li>
        </ol>
        <p>Un <b>dato</b> es una información suelta («este mes se notificaron 12 incidencias»). Un <b>indicador</b> relaciona la información con un objetivo y permite ver tendencias («% de … frente a un estándar del 95 %»).</p>`)}
      <div class="card">
        ${campo(ind, "subproceso", "Subproceso", { ph: "p. ej. Valoración de enfermería al ingreso", help: "La fase del proceso asistencial donde está el riesgo. Sale del Mapa de Riesgos (AMFE) de tu unidad." })}
        ${campo(ind, "hallazgo", "Hallazgo o fallo del Mapa de Riesgos", { area: true, ph: "p. ej. No se valora el riesgo de úlceras por presión al ingreso", help: "Lo que puede fallar, tal como aparece en el Mapa de Riesgos o en una auditoría anterior." })}
        <div class="field">
          <div class="field-head"><label class="field-label">Prioridad en el AMFE</label><span class="field-opt">(opcional)</span>${ayudaBtn("Gravedad (G), ocurrencia (O) y detección (D), de 1 a 10 cada una, como en tu Mapa de Riesgos. IPR = G × O × D (máximo 1.000). Sirve para ordenar el cuadro y empezar por lo más prioritario.").btn}</div>
          <p class="help" id="help-h${nAyuda}" hidden>Gravedad (G), ocurrencia (O) y detección (D), de 1 a 10 cada una, como en tu Mapa de Riesgos. IPR = G × O × D (máximo 1.000). Sirve para ordenar el cuadro y empezar por lo más prioritario.</p>
          <div class="god">
            ${["g", "o", "d"].map(k => `<label class="god-c"><span>${{ g: "Gravedad", o: "Ocurrencia", d: "Detección" }[k]}</span><input class="txt-in" data-k="${k}" value="${esc(ind[k])}" inputmode="numeric" placeholder="1–10" /></label>`).join("")}
            <div class="god-r" data-zona="ipr">${iprHtml(ipr, ind)}</div>
          </div>
        </div>
        ${campo(ind, "criterio", "Criterio (la buena práctica)", { extra: `<div class="quick no-print"><button type="button" class="chip" data-act="ind-lc" title="Escribe el criterio desde la pregunta PICO de un artículo leído con la app Lectura Crítica y guarda su evidencia">📄 Traer de Lectura Crítica</button></div>`, area: true, ph: "p. ej. Se registra la valoración del riesgo de UPP en las primeras 24 horas del ingreso", help: "La regla que debería cumplirse. Escríbela en positivo, con un plazo o condición concreta. Tomada de un protocolo interno, de los estándares ACSA o de una guía." })}
        ${campo(ind, "nombre", "Nombre del indicador", { area: true, ph: "p. ej. % de pacientes ingresados con la valoración del riesgo de UPP registrada en las primeras 24 horas", help: "El criterio convertido en algo que se cuenta: «% de [casos] en los que [criterio]», «Nº de…», «Tasa de… por 1.000…». Que se entienda sin explicaciones.",
          extra: `<div class="quick"><button type="button" class="chip chip--accent" data-act="sugerir-nombre">✎ Proponer nombre a partir del criterio</button></div>` })}
      </div>
      ${ejemplo("del criterio al indicador", `<p><b>Criterio:</b> «se registra la valoración del riesgo de úlceras por presión en las primeras 24 horas del ingreso».</p><p><b>Indicador:</b> «porcentaje de pacientes ingresados con la valoración del riesgo de úlceras por presión registrada en las primeras 24 horas».</p>`)}
      ${trampa("Un dato no es un indicador. «Se notificaron 12 caídas» es un dato; «caídas por cada 1.000 estancias, frente a un estándar» ya es un indicador.")}`;
  }
  function iprHtml(ipr, ind) {
    if (ipr) return `<span class="ipr-val">IPR <b>${ipr}</b></span>`;
    if ([ind.g, ind.o, ind.d].some(lleno)) return `<span class="ipr-falta">G, O y D: enteros de 1 a 10</span>`;
    return `<span class="ipr-falta">IPR = G × O × D</span>`;
  }

  // ----- Paso 2 -----
  const PREGUNTA_TIPO = { estructura: "¿Se dispone de…?", proceso: "¿Se hace…?", resultado: "¿Qué le pasa al paciente?", cumplimiento: "¿Se respeta la norma…?" };
  const EJ_TIPO = { estructura: "Ratio de enfermeras por cama; dispensadores de solución hidroalcohólica en el punto de atención.", proceso: "% de pacientes con la lista de verificación quirúrgica aplicada.", resultado: "Tasa de infección de localización quirúrgica; caídas por 1.000 estancias.", cumplimiento: "Grado de cumplimiento de los estándares ACSA de seguridad del paciente." };
  const EJ_CALC = { centinela: "Cirugía en lugar equivocado", absoluto: "Nº de caídas con daño al mes", porcentaje: "% de pacientes con pulsera identificativa", tasa: "UPP adquiridas por 1.000 estancias", media: "Tiempo medio hasta la primera atención", ratio: "Nº medio de campos sin rellenar por informe de alta", compuesto: "% de cumplimiento del paquete de valoración al ingreso" };
  function paso2(ind) {
    return `
      ${idea("Dos preguntas distintas: <b>¿qué miro?</b> (estructura, proceso, resultado o cumplimiento) y <b>¿cómo lo calculo?</b> (centinela, porcentaje, tasa…). La segunda decide si se puede trabajar con muestra.")}
      ${facil(`
        <h3>¿Qué miro?</h3>
        <ul><li><b>Estructura</b>: lo que se tiene (personal, equipos, instalaciones).</li><li><b>Proceso</b>: lo que se hace (si se aplican los protocolos).</li><li><b>Resultado</b>: lo que se logra (el efecto en el paciente).</li><li><b>Cumplimiento</b>: cómo se respeta lo acordado (normas, estándares).</li></ul>
        <p>Un cuadro equilibrado mezcla los tipos: los de proceso avisan pronto y dicen qué corregir; los de resultado dicen si de verdad mejora el paciente.</p>
        <h3>¿Cómo lo calculo?</h3>
        <ul><li><b>Centinela</b> = un detector de humo: no esperas a que suene tres veces; la primera ya obliga a actuar.</li><li><b>Porcentaje</b> = «de cada 100, ¿cuántos lo hicieron bien?». El más usado en auditoría de historias.</li><li><b>Compuesto</b> = la nota final de la asignatura, hecha con las notas de varios exámenes.</li></ul>`)}
      <div class="card">
        <div class="field"><div class="field-head"><label class="field-label">¿Qué mide? (según lo que miden)</label></div>
          <div class="opt-grid">${Object.keys(L.TIPOS).map(k => { const t = L.TIPOS[k]; return `<button type="button" class="opt-card opt-card--${k}${ind.tipo === k ? " active" : ""}" data-pick="tipo" data-v="${k}"><span class="opt-letra">${t.letra}</span><span class="opt-txt"><b>${t.nom}</b><span class="opt-q">${t.que} · ${PREGUNTA_TIPO[k]}</span><span class="opt-ej">${esc(EJ_TIPO[k])}</span></span></button>`; }).join("")}</div>
        </div>
        ${campo(ind, "justTipo", "¿Por qué ese tipo?", { area: true, ph: "p. ej. Comprueba si se hace una actividad del protocolo, no el resultado final", help: "Explicar por qué has elegido ese tipo te obliga a pensar si de verdad mide lo que crees." })}
        <div class="field"><div class="field-head"><label class="field-label">¿Cómo se calcula? (forma de cálculo)</label></div>
          <div class="opt-grid opt-grid--calc">${Object.keys(L.CALCULOS).map(k => { const c = L.CALCULOS[k]; return `<button type="button" class="opt-card${ind.calculo === k ? " active" : ""}" data-pick="calculo" data-v="${k}"><span class="opt-txt"><b>${c.nom}</b><span class="opt-q">${c.que}</span><span class="opt-ej">${esc(EJ_CALC[k])}</span><span class="opt-muestra opt-muestra--${c.muestra === "sí" ? "si" : "no"}">¿Con muestra? ${c.muestra}</span></span></button>`; }).join("")}</div>
        </div>
        <div data-zona="avisos">${avisosHtml(ind, 2)}</div>
      </div>
      ${ejemplo("conciliación de la medicación", `<p>Una unidad detecta en su Mapa de Riesgos que la conciliación de la medicación al ingreso es muy vulnerable. Define dos indicadores:</p><ul><li><b>Proceso · porcentaje</b>: % de historias con la conciliación realizada y documentada en las primeras 24 h.</li><li><b>Resultado · número absoluto</b>: nº de eventos adversos por omisiones o duplicidades de medicación al alta.</li></ul>`)}
      ${trampa("La media puede maquillar. Un tiempo medio de espera de 20 minutos parece buenísimo, pero puede esconder pacientes de tres horas compensados por otros atendidos al instante. Mejor: «% de pacientes que esperan más de 30 minutos».")}`;
  }

  // ----- Paso 3 -----
  const LBL_NUM = { porcentaje: ["Numerador: casos que cumplen", "p. ej. pacientes con la escala Braden registrada en las primeras 24 h"], tasa: ["Numerador: eventos que se cuentan", "p. ej. caídas con daño notificadas"], media: ["Variable que se mide en cada caso", "p. ej. minutos desde la llegada hasta la primera atención"], ratio: ["Numerador: defectos o cumplimientos que se cuentan", "p. ej. campos sin rellenar"], absoluto: ["Qué se cuenta", "p. ej. caídas con daño"], centinela: ["Suceso centinela", "p. ej. intervención en el lado equivocado"], compuesto: ["Qué se cuenta", "p. ej. criterios del paquete cumplidos"] };
  const LBL_DEN = { porcentaje: ["Denominador: casos evaluados", "p. ej. pacientes ingresados más de 24 h revisados"], tasa: ["Denominador: población y tiempo", "p. ej. estancias del periodo"], media: ["Casos medidos", "p. ej. pacientes atendidos"], ratio: ["Casos revisados", "p. ej. informes de alta revisados"] };
  function paso3(ind) {
    const c = ind.calculo || "porcentaje", num = LBL_NUM[c], den = LBL_DEN[c];
    const sug = L.sentidoSugerido([ind.nombre, ind.numerador].join(" "));
    const unidadPh = L.unidadPorDefecto(ind) || "p. ej. minutos";
    const esPct = ["porcentaje", "compuesto"].includes(ind.calculo);
    return `
      ${idea("La definición operativa es la receta: tiene que quedar tan clara que <b>dos personas distintas, con la misma historia, cuenten lo mismo</b>.")}
      ${facil(`
        <p>Una buena definición operativa recoge, para cada indicador: <b>nombre</b>, <b>fórmula</b>, <b>unidad</b>, <b>fuente de datos</b>, <b>estándar o límite aceptable</b>, <b>periodicidad</b>, <b>responsable</b> y <b>observaciones</b>.</p>
        <h3>Estándar y umbral, sin líos</h3>
        <ul><li><b>Estándar</b> (o meta): el nivel que te comprometes a alcanzar. A veces se llama «umbral» o «límite aceptable». Ejemplo: ≥ 95 %.</li><li><b>Umbral de alarma</b>: el nivel por debajo del cual la situación es claramente mala. Solo hace falta si vas a decidir con LQAS. Ejemplo: 75 %.</li></ul>
        <h3>El sentido</h3>
        <p>Si el indicador cuenta cosas buenas (valoraciones hechas), <b>más es mejor</b> y el estándar va con ≥. Si cuenta fallos (errores, caídas, demoras), <b>menos es mejor</b> y va con ≤.</p>`)}
      ${!ind.calculo ? `<div class="note">ℹ Elige primero la forma de cálculo en el <a href="#/indicador/${ind.id}/2">paso 2</a>: los campos se adaptan a ella.</div>` : ""}
      <div class="card">
        ${campo(ind, "numerador", num[0], { area: true, ph: num[1], help: "Qué casos se cuentan arriba. Si es un porcentaje, son los que cumplen el criterio, de entre los del denominador." })}
        ${den ? campo(ind, "denominador", den[0], { area: true, ph: den[1], help: "Sobre qué casos se calcula. En un porcentaje, todos los casos en los que se podía cumplir el criterio." }) : ""}
        ${c === "tasa" ? campo(ind, "base", "Base de la tasa", { ph: "1.000", quick: ["100", "1.000", "10.000", "100.000"], help: "Por cuántas unidades de población-tiempo se expresa (por 1.000 estancias, por 100 ingresos…)." }) : ""}
        <div class="field"><div class="field-head"><label class="field-label">Fórmula</label><span class="field-opt">(se escribe sola)</span></div><div class="formula" data-zona="formula">${esc(L.formula(ind)) || "Elige la forma de cálculo en el paso 2."}</div></div>
        ${campo(ind, "unidad", "Unidad", { ph: unidadPh, opt: true, help: "Si lo dejas vacío se usa «" + esc(unidadPh || "—") + "»." })}
        ${campo(ind, "poblacion", "Población (a quién se aplica)", { area: true, ph: "p. ej. Pacientes ingresados en la unidad en el trimestre", help: "Qué casos entran: servicio, tipo de paciente, periodo." })}
        ${campo(ind, "exclusiones", "Exclusiones", { area: true, opt: true, ph: "p. ej. Estancias de menos de 24 h", help: "Casos que no cuentan y por qué. Evitan discusiones al auditar." })}
        <div class="field"><div class="field-head"><label class="field-label">Fuente de datos</label>${ayudaBtn("De dónde sale la información. Si es la historia clínica, normalmente habrá que muestrear (paso 5).").btn}</div><p class="help" id="help-h${nAyuda}" hidden>De dónde sale la información. Si es la historia clínica, normalmente habrá que muestrear (paso 5).</p>
          <div class="quick">${L.FUENTES.map(f => `<button type="button" class="chip${(ind.fuentes || []).includes(f) ? " active" : ""}" data-multi="fuentes" data-v="${esc(f)}">${esc(f)}</button>`).join("")}</div>
          <div class="input-wrap" style="margin-top:8px"><input class="txt-in" data-k="fuenteOtra" value="${esc(ind.fuenteOtra)}" placeholder="Otra o más detalle (módulo, registro…)" /></div>
        </div>
        <div class="field"><div class="field-head"><label class="field-label">Sentido</label>${ayudaBtn("¿Qué es bueno: que el valor suba o que baje? Decide si el estándar va con ≥ o con ≤.").btn}</div><p class="help" id="help-h${nAyuda}" hidden>¿Qué es bueno: que el valor suba o que baje? Decide si el estándar va con ≥ o con ≤.</p>
          ${seg(ind, "sentido", [["mas", "▲ Más es mejor (≥)"], ["menos", "▼ Menos es mejor (≤)"]])}
          ${sug && !ind.sentido ? `<p class="field-sub">Por el nombre, parece que <b>${sug === "menos" ? "menos es mejor" : "más es mejor"}</b>.</p>` : ""}
        </div>
        ${c === "centinela" ? `<div class="note">ℹ En un centinela el estándar es siempre <b>0 casos</b>: cualquier caso es un aviso.</div>` : `
        ${campo(ind, "estandar", "Estándar o meta", { num: true, unit: esPct ? "%" : (L.unidad(ind) || ""), ph: esPct ? "p. ej. 95" : "", help: "El nivel que te comprometes a alcanzar (a veces llamado «umbral» o «límite aceptable»). Se compara con él cada medición. Puede salir de un estándar ACSA, de una guía, de tu histórico o de otras unidades (benchmark).", quick: esPct ? (ind.sentido === "menos" ? ["1", "2", "5", "10"] : ["80", "85", "90", "95", "100"]) : null })}
        ${esPct ? campo(ind, "umbral", "Umbral de alarma", { num: true, unit: "%", opt: true, ph: "p. ej. 75", help: "Solo para decidir con LQAS (paso 5): el nivel a partir del cual la situación es claramente mala. Debe quedar " + (ind.sentido === "menos" ? "por encima" : "por debajo") + " del estándar. Cuanto más cerca del estándar, más historias habrá que revisar." }) : ""}`}
        <div data-zona="estandar">${estandarHtml(ind)}</div>
        <div data-zona="avisos">${avisosHtml(ind, 3)}</div>
        ${campo(ind, "periodicidad", "Periodicidad", { ph: "p. ej. Trimestral", quick: L.PERIODICIDADES, help: "Cada cuánto se mide. Hace al indicador «Temporal» (la T de SMART)." })}
        ${campo(ind, "responsable", "Responsable", { ph: "p. ej. Supervisión de enfermería de la unidad", help: "Quién mide y quién puede actuar con el resultado. Sin responsable, el indicador no es útil." })}
        ${campo(ind, "terminos", "Definición de términos", { area: true, rows: 3, destacado: true, ph: "p. ej. «Registrada»: escala completa con su puntuación. «24 horas»: desde la hora de llegada a la unidad.", help: "Aclara cada palabra que se pueda entender de dos maneras: «correctamente», «completo», «a tiempo», «paciente frágil»… Es lo que hace fiable al indicador." })}
        ${campo(ind, "observaciones", "Observaciones", { area: true, opt: true, ph: "p. ej. Se presenta en la Comisión de Calidad cada trimestre", help: "Lo que haya que saber para recoger o interpretar el dato: limitaciones, cambios de registro, a quién se presenta…" })}
      </div>
      ${ejemplo("valoración del riesgo de UPP", `<p><b>Fórmula:</b> (Nº de pacientes con la escala Braden completa registrada en las primeras 24 h ÷ Nº de pacientes ingresados más de 24 h revisados) × 100.</p><p><b>Fuente:</b> historia clínica. <b>Estándar:</b> ≥ 95 %. <b>Periodicidad:</b> trimestral. <b>Responsable:</b> supervisión de la unidad.</p><p><b>Términos:</b> «registrada» = los seis ítems y la puntuación total; «24 horas» = desde la hora de llegada a la unidad.</p>`)}
      ${trampa("El sentido al revés. «% de discrepancias… ≥ 95 %» pide que casi todo sean errores. Si el indicador cuenta fallos, el estándar va con ≤ (por ejemplo, ≤ 5 %); si cuenta aciertos, con ≥.")}`;
  }
  function estandarHtml(ind) {
    const t = L.textoEstandar(ind);
    return t ? `<p class="estandar-prev">Se leerá así: <b>${esc(t)}</b>${lleno(ind.umbral) && L.textoEstandar(ind, "umbral") ? ` · umbral de alarma ${esc(L.textoEstandar(ind, "umbral").replace(/^[≥≤] /, ""))}` : ""}</p>` : "";
  }

  // ----- Paso 4 -----
  const PREG_CAL = {
    E: ["¿Se centra en un aspecto concreto y lo entiende todo el mundo?", "Específico (Specific)"],
    M: ["¿Se puede contar de forma objetiva, sin opiniones?", "Medible (Measurable)"],
    A: ["¿Es viable medirlo con los recursos que hay?", "Alcanzable (Achievable)"],
    R: ["¿Está ligado a un riesgo u objetivo importante y servirá para decidir?", "Relevante (Relevant)"],
    T: ["¿Está fijado cada cuánto se mide?", "Temporal (Time-bound)"],
    valido: ["¿Mide de verdad lo que dice medir? Contar hojas rellenadas no sirve si lo que importa es la valoración bien hecha.", "Válido"],
    fiable: ["¿Dos auditoras que revisan la misma historia llegan a la misma conclusión?", "Fiable"],
    util: ["¿Quien recibe el resultado puede hacer algo para mejorarlo?", "Útil"],
    comparable: ["¿Se puede comparar en el tiempo y entre unidades parecidas (mismo método)?", "Comparable"]
  };
  function paso4(ind) {
    const P = L.pistasCalidad(ind), ac = L.acuerdoPiloto(ind.piloto), links = L.enlacesTM(ind);
    const fila = q => {
      const c = (ind.calidad || {})[q.k] || {}, p = P[q.k];
      return `<div class="cal-row cal-row--${c.v || "nada"}">
        <div class="cal-head"><span class="cal-letra">${q.k.length === 1 ? q.k : "+"}</span><div><b>${PREG_CAL[q.k][1]}</b><p class="cal-q">${PREG_CAL[q.k][0]}</p></div></div>
        <p class="cal-pista cal-pista--${p.ok ? "ok" : "falta"}">${p.ok ? "✓" : "!"} ${esc(p.t)}</p>
        <div class="seg seg--cal">${[["si", "Sí"], ["dudoso", "Dudoso"], ["no", "No"]].map(([v, t]) => `<button type="button" class="${c.v === v ? "active seg-" + v : ""}" data-pick="calidad.${q.k}.v" data-v="${v}">${t}</button>`).join("")}</div>
        <textarea class="txt-in txt-in--sm" data-k="calidad.${q.k}.t" rows="1" placeholder="Cómo se cumple (opcional)">${esc(c.t || "")}</textarea>
      </div>`;
    };
    return `
      ${idea("No todo lo que se puede medir aporta valor. Antes de lanzarlo, pásale el filtro <b>SMART</b> y las tres preguntas: <b>¿válido? ¿fiable? ¿útil?</b>")}
      ${facil(`
        <p><b>SMART</b>: Específico, Medible, Alcanzable, Relevante y Temporal. Algunas plantillas ponen «Realista» en la R; aquí usamos <b>Relevante</b> porque «realista» ya está dentro de «alcanzable».</p>
        <p>Además: <b>válido</b> (mide lo que dice medir), <b>fiable</b> (mide siempre lo mismo en las mismas condiciones), <b>útil</b> (alguien puede actuar con el resultado) y <b>comparable</b> (en el tiempo y entre unidades, con el mismo método).</p>
        <p>Las pistas ✓ / ! salen de lo que has escrito en los pasos anteriores. La valoración la haces tú.</p>`)}
      <div class="card">
        <div data-zona="semaforo">${semaforoCalHtml(ind)}</div>
        <h3 class="sub-h">SMART</h3>
        ${L.CALIDAD.filter(q => q.grupo === "SMART").map(fila).join("")}
        <h3 class="sub-h">Además</h3>
        ${L.CALIDAD.filter(q => q.grupo !== "SMART").map(fila).join("")}
      </div>
      <div class="card piloto">
        <h3 class="sub-h">Pilótalo con dos auditoras</h3>
        <p class="field-sub">Elige al azar unas 30–50 historias (que haya de las que cumplen y de las que no). Dos personas las revisan <b>por separado</b> y se cuenta en cuántas coinciden.</p>
        <div class="piloto-grid">
          ${[["a", "Las dos «cumple»"], ["b", "A «cumple» · B «no»"], ["c", "A «no» · B «cumple»"], ["d", "Las dos «no cumple»"]].map(([k, t]) => `<label class="an-campo"><span>${t}</span><input class="txt-in" data-k="piloto.${k}" value="${esc(ind.piloto[k])}" inputmode="numeric" placeholder="0" /></label>`).join("")}
        </div>
        <div data-zona="piloto">${pilotoHtml(ac)}</div>
        <a class="btn btn--primary btn--tm" data-zona-href="concordancia" href="${esc(links.concordancia)}" target="_blank" rel="noopener">Calcular el kappa en Tamaño Muestral ↗</a>
      </div>
      ${trampa("Si dos auditoras revisan la misma historia y una marca «cumple» y la otra «no cumple», el problema no está en la historia, sino en el indicador, que estaba mal definido. Aclara los términos ambiguos y repite el pilotaje.")}`;
  }
  function semaforoCalHtml(ind) {
    const e = L.estadoCalidad(ind);
    const txt = { gris: "Valora cada característica con Sí, Dudoso o No.", verde: "Todas las características se cumplen: listo para pilotar y usar.", ambar: "Hay características dudosas o sin valorar: revísalas antes de lanzarlo.", rojo: "Alguna característica no se cumple: conviene rehacer esa parte del indicador." }[e.color];
    return `<div class="an-sem an-sem--${e.color === "ambar" ? "gris" : e.color === "gris" ? "nada" : e.color}"><span class="sem-luz"></span><div><b>${e.texto}</b><span>${txt} (${e.contestadas} de ${e.total} valoradas)</span></div></div>`;
  }
  function pilotoHtml(ac) {
    if (!ac) return `<p class="field-sub">Rellena las cuatro casillas para ver el porcentaje de acuerdo.</p>`;
    return `<div class="piloto-res"><b>${nf(ac.acuerdo * 100)} %</b> de acuerdo en ${ac.n} historias (${ac.discrepan} con discrepancia).</div>
      <p class="field-sub">El acuerdo bruto engaña si casi todo cumple: el <b>kappa</b> descuenta el acuerdo que se daría por azar. Revisa juntas las ${ac.discrepan} discrepancias y aclara la definición.</p>`;
  }

  // ----- Paso 5 -----
  function paso5(ind) {
    const rec = L.recomendacionMuestra(ind), links = L.enlacesTM(ind), m = ind.muestra, lp = L.lqasParams(ind);
    const opciones = [];
    if (["porcentaje", "tasa", "ratio", ""].includes(ind.calculo) || rec.modo === "elegir") {
      opciones.push(["estimar", "⚖ Estimar el valor", "La báscula de precisión: «el 78 % cumple, ±10 %». Hacen falta más historias."]);
      if (ind.calculo === "porcentaje") opciones.push(["decidir", "🚦 Decidir si hay problema (LQAS)", "El arco del aeropuerto: no dice cuánto, solo si pita. Con muy pocas historias."]);
    }
    if (ind.calculo === "media") opciones.push(["media", "⚖ Estimar la media", "Cuántos casos para conocer la media con un margen de error."]);
    if (ind.calculo === "compuesto") opciones.push(["compuesto", "Σ Indicador compuesto", "Se revisan pacientes y se cuentan oportunidades de cada criterio."]);
    opciones.push(["sin", "Todos los casos (sin muestra)", ind.calculo === "centinela" ? "Lo obligado en un centinela." : "Si la fuente da todos los casos (sistema de información, notificaciones) o son pocos."]);
    let det = "";
    if (ind.objetivo === "estimar") det = `
      <h3 class="sub-h">Estimar un porcentaje</h3>
      ${campo(ind, "muestra.conf", "Confianza", { num: true, unit: "%", quick: ["90", "95", "99"], help: "Lo habitual es el 95 %." })}
      ${campo(ind, "muestra.e", "Margen de error (±)", { num: true, unit: "%", quick: ["5", "10", "15"], help: "Cuánto te puedes equivocar. ±10 % suele bastar en auditoría interna. Doble de precisión = cuatro veces más historias." })}
      ${campo(ind, "muestra.p", "Porcentaje esperado", { num: true, unit: "%", opt: true, ph: "50", help: "Lo que esperas encontrar. Si no tienes ni idea, deja 50 %: es lo más exigente y vas sobre seguro." })}
      ${campo(ind, "muestra.N", "Tamaño del universo (N)", { num: true, unit: "casos", opt: true, help: "Cuántos casos hay en total para elegir. Si es pequeño, hacen falta menos historias." })}
      ${campo(ind, "muestra.loss", "Pérdidas previstas", { num: true, unit: "%", opt: true, help: "Historias que no se podrán evaluar (no encontradas, incompletas)." })}
      <a class="btn btn--primary btn--tm" data-zona-href="proporcion" href="${esc(links.proporcion)}" target="_blank" rel="noopener">Calcular la muestra en Tamaño Muestral ↗</a>
      ${practica("Regla práctica: en evaluaciones internas se usan unas <b>50 historias al azar</b> (mínimo 30). Con tan pocas, solo se detectan cambios grandes.")}`;
    else if (ind.objetivo === "decidir") det = `
      <h3 class="sub-h">Decidir con LQAS</h3>
      <div class="lq-resumen">
        <div><span>Estándar</span><b>${lp ? nf(lp.pU) + " %" : "—"}</b></div><div><span>Umbral</span><b>${lp ? nf(lp.pL) + " %" : "—"}</b></div>
      </div>
      ${!lleno(ind.estandar) || !lleno(ind.umbral) ? `<div class="warn">⚠ Para el LQAS hacen falta el <b>estándar</b> y el <b>umbral de alarma</b>. Escríbelos aquí o en el <a href="#/indicador/${ind.id}/3">paso 3</a>.</div>
        ${campo(ind, "estandar", "Estándar o meta", { num: true, unit: "%" })}${campo(ind, "umbral", "Umbral de alarma", { num: true, unit: "%" })}` : ""}
      ${lp && lp.invertido ? `<div class="note">ℹ Tu indicador cuenta fallos (menos es mejor). El LQAS trabaja con casos <b>correctos</b>, así que se convierte: ≤ ${nf(100 - lp.pU)} % de fallos = <b>${nf(lp.pU)} %</b> de casos correctos; umbral ${nf(100 - lp.pL)} % de fallos = <b>${nf(lp.pL)} %</b> correctos.</div>` : ""}
      ${campo(ind, "muestra.alpha", "Falsa alarma (α)", { num: true, unit: "%", quick: ["5", "10"], help: "Riesgo de rechazar un servicio que en realidad cumple el estándar (suena el detector y no hay fuego). Lo habitual: 5 %." })}
      ${campo(ind, "muestra.beta", "Fuego sin alarma (β)", { num: true, unit: "%", quick: ["10", "20"], help: "Riesgo de dar por bueno un servicio que está en el umbral o peor. Lo habitual: 20 % o 10 %." })}
      ${campo(ind, "muestra.N", "Tamaño del lote (N)", { num: true, unit: "casos", opt: true, help: "Cuántos casos tiene el lote que quieres juzgar. Si no lo sabes, déjalo vacío." })}
      ${lp ? `<a class="btn btn--primary btn--tm" data-zona-href="lqas" href="${esc(links.lqas)}" target="_blank" rel="noopener">Calcular el plan LQAS en Tamaño Muestral ↗</a>` : ""}
      ${trampa("El LQAS no estima el cumplimiento: si en 16 historias cumplen 12, no puedes escribir «el cumplimiento es del 75 %». Solo «hay» o «no hay evidencia de problema». Y las tablas clásicas de Lemeshow son aproximadas: Tamaño Muestral calcula el plan exacto.")}`;
    else if (ind.objetivo === "media") det = `
      <h3 class="sub-h">Estimar una media</h3>
      ${campo(ind, "muestra.conf", "Confianza", { num: true, unit: "%", quick: ["90", "95", "99"] })}
      ${campo(ind, "muestra.sigma", "Desviación típica (σ)", { num: true, unit: L.unidad(ind) || "", help: "Cuánto varían los casos entre sí. Sácala de un estudio previo, de un piloto o de la literatura." })}
      ${campo(ind, "muestra.eMedia", "Margen de error (±)", { num: true, unit: L.unidad(ind) || "", help: "En las mismas unidades que la variable (p. ej. ±5 minutos)." })}
      ${campo(ind, "muestra.N", "Tamaño de la población (N)", { num: true, unit: "casos", opt: true })}
      <a class="btn btn--primary btn--tm" data-zona-href="media" href="${esc(links.media)}" target="_blank" rel="noopener">Calcular la muestra en Tamaño Muestral ↗</a>`;
    else if (ind.objetivo === "compuesto") det = `
      <h3 class="sub-h">Indicador compuesto</h3>
      <p>Monta el compuesto con su desglose en el <a href="#/compuesto/nuevo">módulo de compuestos</a>. Para su análisis estadístico, Tamaño Muestral tiene una <a href="${L.TM_URL}#/compuesto" target="_blank" rel="noopener">calculadora de compuestos ↗</a>.</p>`;
    else if (ind.objetivo === "sin") det = `
      <h3 class="sub-h">Sin muestra</h3>
      <p>${ind.calculo === "centinela" ? "Se vigilan <b>todos los casos</b>. Cada caso se notifica y se analiza (análisis de causas) aunque sea uno solo." : "Se cuentan <b>todos los casos</b> del periodo. Comprueba que la fuente los recoge todos (la infranotificación hace que el número parezca mejor de lo que es)."}</p>`;
    return `
      ${idea("Cuando la fuente son historias clínicas rara vez se pueden revisar todas. Antes de calcular, decide qué quieres: <b>estimar</b> cuánto cumples o <b>decidir</b> si hay un problema.")}
      ${facil(`<p>Esta app no calcula muestras: abre <b>Tamaño Muestral</b> con los datos de tu indicador ya puestos. Allí tienes el desarrollo paso a paso, el plan exacto y el sorteo de las historias.</p><p>Una muestra tiene que ser <b>al azar</b>: «las 20 historias que están sobre la mesa» o «las últimas que se cerraron» suelen dar un resultado más bonito que la realidad.</p>`)}
      <div class="card">
        <div class="note">ℹ ${esc(rec.texto)}</div>
        <div class="opt-grid opt-grid--obj">${opciones.map(([v, t, s]) => `<button type="button" class="opt-card${ind.objetivo === v ? " active" : ""}" data-pick="objetivo" data-v="${v}"><span class="opt-txt"><b>${t}</b><span class="opt-q">${s}</span></span></button>`).join("")}</div>
        ${det}
      </div>
      ${ind.objetivo && ind.objetivo !== "sin" ? `<div class="card">
        <h3 class="sub-h">Anota lo que te dé Tamaño Muestral</h3>
        <p class="field-sub">Así queda en la ficha y en el cuadro.</p>
        <div class="plan-grid">
          ${campo(ind, "plan.n", "Casos a revisar (n)", { num: true })}
          ${ind.objetivo === "decidir" ? campo(ind, "plan.c", "Número de decisión (c)", { num: true, help: "Mínimo de casos que deben cumplir en la muestra para aceptar el lote." }) : ""}
        </div>
        ${campo(ind, "plan.nota", "Nota", { area: true, opt: true, ph: "p. ej. Muestreo sistemático sobre el listado de ingresos" })}
      </div>` : ""}`;
  }

  // ----- Paso 6 · Ficha -----
  function fichaHtml(ind, o = {}) {
    const t = L.TIPOS[ind.tipo], c = L.CALCULOS[ind.calculo], ipr = L.ipr(ind.g, ind.o, ind.d), est = L.estadoCalidad(ind), ac = L.acuerdoPiloto(ind.piloto);
    const fila = (k, v, paso) => `<tr><th>${k}</th><td>${v || `<span class="vacio">—${paso && !o.soloLectura ? ` <a class="no-print" href="#/indicador/${ind.id}/${paso}">completar</a>` : ""}</span>`}</td></tr>`;
    const cal = L.CALIDAD.map(q => { const x = (ind.calidad || {})[q.k] || {}; return `<tr><th>${q.nom}</th><td><span class="cal-chip cal-chip--${x.v || "nada"}">${{ si: "Sí", dudoso: "Dudoso", no: "No" }[x.v] || "Sin valorar"}</span>${lleno(x.t) ? " " + esc(x.t) : ""}</td></tr>`; }).join("");
    return `<article class="ficha">
      <div class="ficha-cab">
        <div class="badges">${t ? `<span class="badge badge--${ind.tipo}">${t.nom}</span>` : ""}${c ? `<span class="badge">${c.nom}</span>` : ""}${ipr ? `<span class="badge badge--ipr">IPR ${ipr}</span>` : ""}<span class="badge badge--sem badge--${est.color}"><span class="punto"></span>Calidad: ${est.texto}</span></div>
        <h2 class="ficha-nombre">${lleno(ind.nombre) ? esc(ind.nombre) : '<span class="vacio">Indicador sin nombre</span>'}</h2>
      </div>
      <section class="ficha-sec"><h3>1 · Origen</h3><table class="ficha-t">
        ${fila("Subproceso", esc(ind.subproceso), 1)}${fila("Hallazgo", escMl(ind.hallazgo), 1)}${ipr ? fila("Prioridad (AMFE)", `G ${esc(ind.g)} × O ${esc(ind.o)} × D ${esc(ind.d)} = <b>IPR ${ipr}</b>`) : ""}${fila("Criterio", escMl(ind.criterio), 1)}
      </table></section>
      <section class="ficha-sec"><h3>2 · Caracterización</h3><table class="ficha-t">
        ${fila("Qué mide", t ? `${t.nom} <span class="gris">(${t.que.toLowerCase()})</span>` : "", 2)}${fila("Justificación", escMl(ind.justTipo), 2)}${fila("Forma de cálculo", c ? c.nom : "", 2)}
      </table></section>
      <section class="ficha-sec"><h3>3 · Definición operativa</h3><table class="ficha-t">
        ${fila(ind.calculo === "media" ? "Variable" : "Numerador", escMl(ind.numerador), 3)}
        ${L.usaDenominador(ind.calculo) ? fila("Denominador", escMl(ind.denominador), 3) : ""}
        ${fila("Fórmula", ind.calculo ? `<span class="formula-mini">${esc(L.formula(ind))}</span>` : "", 3)}
        ${fila("Unidad", esc(L.unidad(ind)))}
        ${fila("Población", escMl(ind.poblacion), 3)}
        ${lleno(ind.exclusiones) ? fila("Exclusiones", escMl(ind.exclusiones)) : ""}
        ${fila("Fuente de datos", esc(L.fuentesTexto(ind)), 3)}
        ${fila("Estándar", `<b>${esc(L.textoEstandar(ind))}</b>${ind.sentido ? ` <span class="gris">(${ind.sentido === "menos" ? "menos es mejor" : "más es mejor"})</span>` : ""}`, 3)}
        ${lleno(ind.umbral) ? fila("Umbral de alarma", esc(L.textoEstandar(ind, "umbral"))) : ""}
        ${fila("Periodicidad", esc(ind.periodicidad), 3)}${fila("Responsable", esc(ind.responsable), 3)}
        ${fila("Definición de términos", escMl(ind.terminos), 3)}
        ${lleno(ind.observaciones) ? fila("Observaciones", escMl(ind.observaciones)) : ""}
      </table></section>
      <section class="ficha-sec"><h3>4 · Calidad del indicador</h3><table class="ficha-t">${cal}
        ${ac ? fila("Pilotaje", `${nf(ac.acuerdo * 100)} % de acuerdo entre dos auditoras en ${ac.n} historias`) : ""}
      </table></section>
      <section class="ficha-sec"><h3>5 · Medición y muestra</h3><table class="ficha-t">
        ${fila("Plan", esc(L.textoPlan(ind)), 5)}
      </table></section>
      ${fichaResultadosHtml(ind, o)}
    </article>`;
  }
  function fichaResultadosHtml(ind, o = {}) {
    const an = L.analizarResultados(ind), tabla = tablaResultadosHtml(ind, an);
    if (!tabla) return o.soloLectura ? "" : `<section class="ficha-sec no-print"><h3>6 · Resultados (Verificar)</h3><p class="field-sub">Aún no hay resultados. <a href="#/indicador/${ind.id}/resultados">Registrar el primer periodo</a></p></section>`;
    return `<section class="ficha-sec ficha-res"><h3>6 · Resultados (Verificar)</h3>
      ${graficoResultados(ind, an)}${tabla}
      ${an.frases.length ? `<ul class="res-frases">${an.frases.map(t => `<li>${esc(t)}</li>`).join("")}</ul>` : ""}
      ${L.esLqas(ind) ? `<p class="field-sub">LQAS: regla de ${esc(ind.plan.c)} correctos de ${esc(ind.plan.n)}. El porcentaje de la muestra no es una medida del cumplimiento.</p>` : ""}
      ${o.soloLectura ? "" : `<p class="no-print"><a href="#/indicador/${ind.id}/resultados">Añadir o cambiar resultados</a></p>`}
    </section>`;
  }
  function paso6(ind) {
    const A = L.avisos(ind).filter(a => a.nivel !== "nota");
    const comp = L.pasosCompletos(ind), faltan = [1, 2, 3, 4, 5].filter(p => !comp[p]);
    return `
      <div class="toolbar no-print">
        <button class="btn btn--primary" data-act="informe">Ver informe</button>
        <a class="btn" href="#/indicador/${ind.id}/resultados">Resultados</a>
        <button class="btn" data-act="imprimir">Guardar / imprimir</button>
        <button class="btn" data-act="enlace">Compartir enlace</button>
        ${modoEjemplo ? `<button class="btn" data-act="copiar-a-mio">Copiar a mi cuadro</button>` : `<button class="btn" data-act="duplicar">Duplicar</button>
        <button class="btn btn--peligro-ghost" data-act="borrar">Borrar</button>`}
      </div>
      ${faltan.length ? `<div class="note no-print">ℹ Pasos sin completar: ${faltan.map(p => `<a href="#/indicador/${ind.id}/${p}">${p} · ${PASOS[p - 1].corto}</a>`).join(", ")}.</div>` : ""}
      ${A.length ? `<div class="no-print">${A.map(a => `<div class="warn">⚠ ${esc(a.texto)}</div>`).join("")}</div>` : ""}
      ${informeCab("Ficha del indicador")}
      ${fichaHtml(ind)}
      ${informePie()}`;
  }

  // ---------- Fase 2 · Resultados por periodo (Verificar del PDCA) ----------
  const ESTADO_TXT = (ind, f) => f.error ? "Revisar" : f.vacio ? "Sin datos" : L.esLqas(ind)
    ? { verde: "Verde", rojo: "Rojo", gris: "Pendiente" }[f.estado] : { verde: "Cumple", rojo: "No cumple", gris: "Sin valorar" }[f.estado];
  const recorta = (t, n) => t.length > n ? t.slice(0, n - 1) + "…" : t;
  // Gráfico SVG sin librerías: línea con la raya del estándar o, en LQAS, una tira de semáforos
  function graficoResultados(ind, an) {
    const pts = an.filas.filter(f => !f.vacio && !f.error && f.valor !== null);
    if (!pts.length) return "";
    const W = 400, H = 210, ml = 44, mr = 14, mt = 16, mb = 38, iw = W - ml - mr, ih = H - mt - mb;
    const x = k => pts.length === 1 ? ml + iw / 2 : ml + k * iw / (pts.length - 1);
    // Etiquetas de los periodos: si no caben en una fila, se alternan en dos alturas
    const cada = Math.ceil(pts.length / 12);
    const ancla = k => pts.length === 1 ? "middle" : k === 0 ? "start" : k === pts.length - 1 ? "end" : "middle";
    const anchoTxt = Math.max(...pts.map(f => recorta(f.periodo, 10).length)) * 7.2;
    const hueco = pts.length > 1 ? iw / (pts.length - 1) * cada : iw;
    const alterna = hueco < anchoTxt * 1.5 + 8;   // la primera y la última van ancladas al borde y ocupan más
    const dyFila = k => alterna && Math.round(k / cada) % 2 ? 13 : 0;
    const etiqueta = (f, k, dx, y0) => k % cada === 0 || k === pts.length - 1 ? `<text class="gr-txt" x="${x(k) + (ancla(k) === "start" ? -dx : ancla(k) === "end" ? dx : 0)}" y="${y0 + dyFila(k)}" text-anchor="${ancla(k)}">${esc(recorta(f.periodo, 10))}</text>` : "";
    const etiquetas = pts.map((f, k) => etiqueta(f, k, 6, H - mb + 16)).join("");
    const titulo = `Resultados de «${esc(ind.nombre || "indicador")}» por periodo`;
    if (L.esLqas(ind)) {
      const y = mt + 40;
      return `<svg class="grafico" viewBox="0 0 ${W} ${mt + 100}" role="img" aria-label="${titulo}"><title>${titulo}</title>
        <line class="gr-eje" x1="${ml - 20}" x2="${W - mr + 6}" y1="${y}" y2="${y}" />
        ${pts.map((f, k) => `<circle class="gr-p gr-p--${f.estado}" cx="${x(k)}" cy="${y}" r="13" /><text class="gr-sem" x="${x(k)}" y="${y + 4.5}" text-anchor="middle">${f.estado === "verde" ? "V" : f.estado === "rojo" ? "R" : "?"}</text>`).join("")}
        ${pts.map((f, k) => etiqueta(f, k, 13, y + 34)).join("")}
        <text class="gr-txt" x="${ml - 20}" y="${mt}">LQAS · regla ${esc(ind.plan.c)} de ${esc(ind.plan.n)} correctos</text>
      </svg>`;
    }
    const est = L.num(ind.estandar), esPct = ["porcentaje", "compuesto"].includes(ind.calculo) || L.unidad(ind) === "%";
    const ej = L.ejeY(pts.map(f => f.valor), ind.calculo === "centinela" ? 0 : est, esPct);
    const y = v => mt + ih - (v - ej.min) / (ej.max - ej.min) * ih;
    const fmt = v => nf(v, 2);
    const marcas = ej.marcas.map(m => `<line class="gr-rej" x1="${ml}" x2="${W - mr}" y1="${y(m)}" y2="${y(m)}" /><text class="gr-txt" x="${ml - 6}" y="${y(m) + 3.5}" text-anchor="end">${fmt(m)}</text>`).join("");
    const linEst = est !== null && !Number.isNaN(est) && est >= ej.min && est <= ej.max
      ? `<line class="gr-est" x1="${ml}" x2="${W - mr}" y1="${y(est)}" y2="${y(est)}" /><text class="gr-est-t" x="${W - mr}" y="${y(est) - 5}" text-anchor="end">Estándar ${ind.sentido === "menos" ? "≤" : ind.sentido === "mas" ? "≥" : ""} ${nf(est, 2)}${L.unidad(ind) === "%" ? " %" : ""}</text>` : "";
    const camino = pts.map((f, k) => `${k ? "L" : "M"}${x(k).toFixed(1)},${y(f.valor).toFixed(1)}`).join(" ");
    return `<svg class="grafico" viewBox="0 0 ${W} ${H}" role="img" aria-label="${titulo}"><title>${titulo}</title>
      ${marcas}${linEst}
      ${pts.length > 1 ? `<path class="gr-linea" d="${camino}" />` : ""}
      ${pts.map((f, k) => `<circle class="gr-p gr-p--${f.estado}" cx="${x(k).toFixed(1)}" cy="${y(f.valor).toFixed(1)}" r="5.5"><title>${esc(f.periodo)}: ${esc(L.textoValor(ind, f.valor))}</title></circle>`).join("")}
      ${etiquetas}
    </svg>`;
  }
  function graficoCardHtml(ind) {
    const an = L.analizarResultados(ind), g = graficoResultados(ind, an);
    if (!g) return `<div class="placeholder-box">Cuando escribas el primer periodo aparecerá aquí el gráfico con la línea del estándar.</div>`;
    return `<div class="chart-card"><h3>${L.esLqas(ind) ? "Semáforo de cada periodo" : "Evolución frente al estándar"}</h3>${g}
      ${L.esLqas(ind) ? `<p>V = verde (sin evidencia de problema) · R = rojo (hay que investigar). LQAS decide; no mide: por eso no se dibuja una línea de porcentajes.</p>` : `<p>La raya discontinua es el estándar. Verde: cumple; rojo: no cumple.</p>`}</div>
      ${an.frases.length ? `<div class="res-resumen"><span class="box-tag">Qué dice la serie</span><ul>${an.frases.map(t => `<li>${esc(t)}</li>`).join("")}</ul></div>` : ""}`;
  }
  // Tabla de solo lectura (ficha, informe y cuadro)
  function tablaResultadosHtml(ind, an) {
    const filas = an.filas.filter(f => !f.vacio);
    if (!filas.length) return "";
    return `<div class="table-wrap"><table class="mini res-t"><thead><tr><th>Periodo</th><th>Datos</th><th>Valor</th><th>Resultado</th><th>Nota</th></tr></thead><tbody>
      ${filas.map(f => { const m = ind.resultados[f.i]; const datosTxt = [m.num, m.den || (f.lqas ? String(f.lqas.rev) : "")].filter(lleno).join(" / ");
        return `<tr><td>${esc(f.periodo)}</td><td>${esc(datosTxt)}</td><td>${esc(L.textoResultado(ind, f))}</td><td><span class="res-est res-est--${f.error ? "error" : f.estado}"><span class="punto"></span>${ESTADO_TXT(ind, f)}</span></td><td>${escMl(m.nota)}</td></tr>`; }).join("")}
    </tbody></table></div>`;
  }
  // Resultado en vivo de una fila del editor
  function filaOutHtml(ind, m, i) {
    const f = L.evaluarMedicion(ind, m);
    if (f.vacio) return "";
    if (f.error) return `<div class="res-out res-out--error">✕ ${esc(f.error)}</div>`;
    const url = L.enlaceAnalisis(ind, m);
    const nota = L.esLqas(ind)
      ? `<span class="res-mini">No escribas «el cumplimiento es del ${nf(f.valor, 0)} %»: con ${f.lqas.rev} casos LQAS decide, no mide.</span>`
      : L.conMuestra(ind) ? `<span class="res-mini">Dato de una muestra: tiene margen de error.</span>` : "";
    return `<div class="res-out res-out--${f.estado}"><span class="punto"></span><div>${L.esLqas(ind) ? `<b>${ESTADO_TXT(ind, f)}</b>` : `<b>${esc(L.textoValor(ind, f.valor))}</b> · ${ESTADO_TXT(ind, f)}`}${/^(No c|C)umple el estándar\.$/.test(f.motivo) ? "" : `<span class="res-mini">${esc(f.motivo)}</span>`}${nota}
      ${url ? `<a class="res-tm" href="${url}" target="_blank" rel="noopener">${L.esLqas(ind) ? "Frase para el informe en Tamaño Muestral ↗" : "Ver la horquilla en Tamaño Muestral ↗"}</a>` : ""}</div></div>`;
  }
  function filaEditorHtml(ind, m, i) {
    const C = L.camposMedicion(ind), plan = ind.plan || {};
    const inp = (k, lab, ph, modo) => `<label>${lab}<input class="txt-in" data-k="resultados.${i}.${k}" value="${esc(m[k])}" placeholder="${esc(ph || "")}" inputmode="${modo || "decimal"}" autocomplete="off" /></label>`;
    return `<div class="res-row">
      <div class="res-row-head"><input class="res-per txt-in" data-k="resultados.${i}.periodo" value="${esc(m.periodo)}" placeholder="${esc(L.ejemploPeriodo(ind.periodicidad))}" aria-label="Periodo" autocomplete="off" />
        <button class="icon-btn" data-act="res-del" data-i="${i}" aria-label="Quitar este periodo">✕</button></div>
      <div class="res-nums">${C.num ? inp("num", C.num, "", "numeric") : ""}${C.den ? inp("den", C.den, L.esLqas(ind) ? plan.n || "" : "", "numeric") : ""}${C.valor ? inp("valor", C.valor, C.valorOblig ? "" : "opc.") : ""}</div>
      <input class="txt-in res-nota" data-k="resultados.${i}.nota" value="${esc(m.nota)}" placeholder="Nota (opcional): incidencias, qué se cambió…" aria-label="Nota" autocomplete="off" />
      <div data-zona="resfila" data-i="${i}" aria-live="polite">${filaOutHtml(ind, m, i)}</div>
    </div>`;
  }
  function renderResultados(ind) {
    actual = ind; pasoActual = 7; nAyuda = 0;
    if (!Array.isArray(ind.resultados)) ind.resultados = [];
    const plan = ind.plan || {}, lqas = L.esLqas(ind);
    const falta = !ind.calculo ? "Elige la forma de cálculo en el paso 2." : ind.calculo !== "centinela" && (!ind.sentido || !lleno(ind.estandar)) ? "Sin sentido (≥ / ≤) y estándar (paso 3) no se puede saber si cada periodo cumple." : lqas && (!lleno(plan.n) || !lleno(plan.c)) ? "Anota n y c del plan LQAS (paso 5) para aplicar la regla." : "";
    view.innerHTML = `
      <div class="wiz-top no-print">
        ${volverCuadro()}
        ${stepperHtml(ind, 7)}
      </div>
      <div class="wiz-titulo"><span class="wiz-paso">Verificar (PDCA)</span><h1 class="page-title">Resultados por periodo</h1></div>
      <p class="page-lead"><b>${lleno(ind.nombre) ? esc(ind.nombre) : "Indicador sin nombre"}</b>${ind.calculo ? ` · <span class="formula-mini">${esc(L.formula(ind))}</span>` : ""}${lleno(ind.estandar) ? ` · estándar <b>${esc(L.textoEstandar(ind))}</b>` : ""}${lqas && lleno(plan.n) ? ` · LQAS n = ${esc(plan.n)}, c = ${esc(plan.c)}` : ""}</p>
      ${falta ? `<div class="warn">⚠ ${esc(falta)} <a href="#/indicador/${ind.id}/${!ind.calculo ? 2 : falta.includes("paso 5") ? 5 : 3}">Completar</a></div>` : ""}
      ${idea("Medir no basta: en la fase <b>Verificar</b> se compara cada periodo con el estándar. Si sale rojo, toca <b>Actuar</b>: analizar causas, mejorar y volver a medir.")}
      ${facil(`<ul>
        <li>Escribe <b>un periodo por fila</b>, en orden (el más antiguo arriba), con la misma periodicidad que en la ficha (${esc(ind.periodicidad || "sin indicar")}).</li>
        <li>${lqas ? `Es un indicador <b>LQAS</b>: escribe el numerador y cuántos casos revisaste. La app aplica la regla que anotaste (verde si hay al menos ${esc(plan.c || "c")} correctos de ${esc(plan.n || "n")}). No da un porcentaje fiable: para la frase del informe, abre «Ya tengo los resultados» en Tamaño Muestral.` : ind.calculo === "centinela" ? "Es un <b>centinela</b>: cualquier caso sale en rojo y se analiza uno a uno." : "Escribe el numerador y el denominador; si ya tienes el valor calculado (p. ej. del sistema de información), puedes escribir solo el valor."}</li>
        <li>Un periodo malo puede ser casualidad; <b>dos o más seguidos</b> son una señal clara. Con menos de tres periodos todavía no hay tendencia.</li>
        ${L.conMuestra(ind) && !lqas ? "<li>Si mides con una <b>muestra</b>, el dato tiene margen de error: mira la horquilla en Tamaño Muestral antes de decir que «ha bajado» o «ha subido».</li>" : ""}
      </ul>`)}
      <div class="calc-grid">
        <div class="card no-print">
          <h2 class="sub-h">Mediciones</h2>
          ${ind.resultados.length ? ind.resultados.map((m, i) => filaEditorHtml(ind, m, i)).join("") : `<p class="field-sub">Todavía no hay ningún periodo.</p>`}
          <div class="toolbar"><button class="btn btn--primary" data-act="res-add">+ Añadir periodo</button></div>
        </div>
        <div class="calc-result-col" data-zona="resgraf">${graficoCardHtml(ind)}</div>
      </div>
      <div class="toolbar no-print">
        <a class="btn" href="#/indicador/${ind.id}/6">Ver la ficha</a>
        <button class="btn" data-act="imprimir">Guardar / imprimir</button>
      </div>`;
  }

  // ---------- Informe (cabecera y pie, solo al imprimir / Ver informe) ----------
  function informeBar() {
    return `<div class="informe-bar informe-only"><button class="btn" data-act="volver">← Volver</button><button class="btn btn--primary" data-act="imprimir">Guardar / imprimir</button></div>`;
  }
  function informeCab(titulo) {
    return `<header class="print-head print-only"><h1>${esc(titulo)}</h1><p>Indicadores · Calidad y seguridad del paciente · ${new Date().toLocaleDateString("es-ES")}</p></header>`;
  }
  function informePie() {
    return `<footer class="print-foot print-only">
      <span class="autor-titulo">Aplicación desarrollada por</span>
      <span class="autor"><b>Carlos J. Galán Doval</b> · Médico de Familia · Técnico en Calidad y Seguridad del Paciente</span>
      <span class="autor"><b>Rafael J. Romero de Castilla Gil</b> · Médico de Familia · Responsable del Área de Calidad y Seguridad del Paciente</span>
      <span class="autor-area">Área de Gestión Sanitaria Sur de Córdoba · Andalucía</span>
      <span class="autor-uso">De libre uso con fines docentes y divulgativos. No sustituye el criterio profesional.</span>
      <span class="footer-version">Indicadores · versión ${APP_VERSION}</span>
    </footer>`;
  }

  // ---------- Render del asistente ----------
  const RENDER_PASO = { 1: paso1, 2: paso2, 3: paso3, 4: paso4, 5: paso5, 6: paso6 };
  function renderWizard(ind, paso) {
    actual = ind; pasoActual = paso; nAyuda = 0;
    const p = PASOS[paso - 1], informe = document.body.classList.contains("ver-informe");
    const ejBtn = esVacio(ind) && paso === 1 ? `<a class="btn" href="#/indicador/ej-1/1">Ver un ejemplo</a>` : "";
    view.innerHTML = `
      ${informe ? informeBar() : ""}
      <div class="wiz-top no-print">
        ${volverCuadro()}
        ${stepperHtml(ind, paso)}
      </div>
      <div class="wiz-grid${paso === 6 ? " wiz-grid--ficha" : ""}">
        <div class="wiz-main">
          <div class="wiz-titulo no-print"><span class="wiz-paso">Paso ${paso} de 6</span><h1 class="page-title">${p.titulo}</h1>${ejBtn}</div>
          ${RENDER_PASO[paso](ind)}
          <div class="wiz-nav no-print">
            ${paso > 1 ? `<a class="btn" href="#/indicador/${ind.id}/${paso - 1}">← Anterior</a>` : `<span></span>`}
            ${paso < 6 ? `<a class="btn btn--primary" href="#/indicador/${ind.id}/${paso + 1}">Siguiente: ${PASOS[paso].corto} →</a>` : `<a class="btn btn--primary" href="#/indicador/${ind.id}/resultados">Registrar resultados →</a>`}
          </div>
        </div>
        ${paso < 6 ? `<aside class="wiz-side no-print" data-zona="resumen">${resumenHtml(ind)}</aside>` : ""}
      </div>`;
  }
  const esVacio = ind => !["subproceso", "hallazgo", "criterio", "nombre", "numerador", "tipo", "calculo"].some(k => lleno(ind[k]));
  // Refresca las zonas que dependen de lo escrito sin repintar los campos (no se pierde el foco)
  function refrescarZonas() {
    if (!actual) return;
    const ind = actual, Z = {
      ipr: () => iprHtml(L.ipr(ind.g, ind.o, ind.d), ind),
      avisos: () => avisosHtml(ind, pasoActual),
      formula: () => esc(L.formula(ind)) || "Elige la forma de cálculo en el paso 2.",
      estandar: () => estandarHtml(ind),
      resumen: () => resumenHtml(ind),
      semaforo: () => semaforoCalHtml(ind),
      piloto: () => pilotoHtml(L.acuerdoPiloto(ind.piloto)),
      resfila: el => filaOutHtml(ind, ind.resultados[Number(el.dataset.i)] || {}, Number(el.dataset.i)),
      resgraf: () => graficoCardHtml(ind)
    };
    $$("[data-zona]").forEach(el => { const f = Z[el.dataset.zona]; if (f) el.innerHTML = f(el); });
    const links = L.enlacesTM(ind);
    $$("[data-zona-href]").forEach(a => { const h = links[a.dataset.zonaHref]; if (h) a.href = h; });
    const st = $(".wiz-steps"); if (st) st.outerHTML = stepperHtml(ind, pasoActual);
  }

  // ---------- Inicio ----------
  function renderHome() {
    const n = datos.indicadores.length, nc = datos.compuestos.length;
    view.innerHTML = `
      <section class="hero">
        <h1 class="hero-t"><span>Indicadores</span></h1>
        <p class="hero-p">Crea, define y comprueba indicadores de calidad y seguridad del paciente paso a paso: del riesgo detectado a la ficha del indicador. Sin necesidad de saber estadística.</p>
        <div class="hero-btns">
          <a class="btn btn--primary btn--grande" href="#/nuevo">+ Nuevo indicador</a>
          <a class="btn btn--grande" href="#/cuadro">Mi cuadro${n || nc ? ` <span class="contador">${n + nc}</span>` : ""}</a>
        </div>
      </section>
      <h2 class="sec-t">El camino en seis pasos</h2>
      <ol class="ruta">
        ${[["Riesgo", "Del hallazgo del Mapa de Riesgos al criterio y al indicador."], ["Tipo", "¿Qué mide? ¿Cómo se calcula?"], ["Definición", "Numerador, denominador, fuente, estándar…"], ["Calidad", "SMART, válido, fiable y útil. Pilotaje con dos auditoras."], ["Muestra", "Estimar o decidir, con Tamaño Muestral."], ["Ficha", "Lista para imprimir, compartir o pasar a Excel. Después, <b>registra los resultados</b> de cada periodo frente al estándar (Verificar)."]].map(([t, s], i) => `<li class="ruta-p"><span class="ruta-n">${i + 1}</span><div><b>${t}</b><span>${s}</span></div></li>`).join("")}
      </ol>
      <div class="guide">
        <a class="guide-card" href="#/compuesto/nuevo"><span class="guide-num">Σ</span><div><div class="guide-q">Indicador compuesto</div><div class="guide-a">Varios criterios en un número, <b>sin perder el desglose</b>.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card" href="#/aprende"><span class="guide-num">?</span><div><div class="guide-q">Aprende</div><div class="guide-a">Tipos, formas de cálculo, SMART, compuestos y su relación con el <b>PDCA</b>.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card" href="#/prueba"><span class="guide-num">✓</span><div><div class="guide-q">Ponte a prueba</div><div class="guide-a">Doce preguntas para comprobar si lo tienes claro.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card" href="#/glosario"><span class="guide-num">A–Z</span><div><div class="guide-q">Glosario</div><div class="guide-a">Cada término <b>en lenguaje llano</b>, con buscador.</div></div><span class="guide-go">›</span></a>
        <!-- [eco] -->
        <a class="guide-card guide-card--hermana" href="${L.TM_URL}" target="_blank" rel="noopener"><span class="guide-num">n</span><div><div class="guide-q">Tamaño Muestral ↗</div><div class="guide-a">La app hermana: <b>cuántos casos revisar</b>, plan LQAS y concordancia entre auditoras.</div></div><span class="guide-go">›</span></a>
        <!-- [/eco] -->
        <!-- [ind]
        <a class="guide-card guide-card--hermana" href="${L.TM_URL}" target="_blank" rel="noopener"><span class="guide-num">n</span><div><div class="guide-q">Tamaño Muestral ↗</div><div class="guide-a">Calculadora incluida: <b>cuántos casos revisar</b>, plan LQAS y concordancia entre auditoras.</div></div><span class="guide-go">›</span></a>
        [/ind] -->
        <!-- [eco] -->
        <a class="guide-card guide-card--hermana" href="${L.AUD_URL}" target="_blank" rel="noopener"><span class="guide-num">↻</span><div><div class="guide-q">Auditorías ↗</div><div class="guide-a">La app hermana: <b>planifica, realiza e informa una auditoría completa</b>, del Mapa de Riesgos al plan de mejora y el seguimiento.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.AMFE_URL}" target="_blank" rel="noopener"><span class="guide-num">IPR</span><div><div class="guide-q">AMFE ↗</div><div class="guide-a">La app hermana: <b>analiza los fallos de un proceso</b> con G·O·D e IPR y elige cuáles medir; desde ella puedes traer tus indicadores a sus acciones de mejora.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.FLUJO_URL}" target="_blank" rel="noopener"><span class="guide-num">→</span><div><div class="guide-q">Flujograma ↗</div><div class="guide-a">La app hermana: <b>dibuja el flujograma de un proceso</b> y marca sus puntos críticos, donde conviene poner un indicador.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.DAFO_URL}" target="_blank" rel="noopener"><span class="guide-num">D</span><div><div class="guide-q">DAFO · CAME ↗</div><div class="guide-a">La app hermana: <b>analiza tu unidad</b> (debilidades, amenazas, fortalezas y oportunidades) y conviértelo en un plan de acción con indicadores.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.MAPA_URL}" target="_blank" rel="noopener"><span class="guide-num">◆</span><div><div class="guide-q">Mapa de Riesgos ↗</div><div class="guide-a">La app hermana: <b>elabora el mapa de riesgos de un proceso</b>, del flujograma y el AMFE a las acciones de mejora; puede traer tus indicadores para medirlas.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.RONDAS_URL}" target="_blank" rel="noopener"><span class="guide-num">✓</span><div><div class="guide-q">Rondas de Seguridad ↗</div><div class="guide-a">La app hermana: <b>rondas de seguridad</b> en una unidad, con su cumplimiento ronda a ronda y las acciones con su indicador.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.LC_URL}" target="_blank" rel="noopener"><span class="guide-num">LC</span><div><div class="guide-q">Lectura Crítica ↗</div><div class="guide-a">App hermana: <b>lee críticamente un artículo</b> para comprobar si el artículo que justifica un indicador o su estándar es fiable.</div></div><span class="guide-go">›</span></a>
        <a class="guide-card guide-card--hermana" href="${L.ACR_URL}" target="_blank" rel="noopener"><span class="guide-num">ACR</span><div><div class="guide-q">ACR · Ishikawa ↗</div><div class="guide-a">App hermana: <b>busca las causas de fondo</b> de un problema o de un incidente (Ishikawa y causa raíz); sus acciones de mejora llevan indicador, que puedes definir aquí.</div></div><span class="guide-go">›</span></a>
        <!-- [/eco] -->
      </div>
      <h2 class="sec-t">Ejemplos resueltos</h2>
      <p class="page-lead">Tócalos para verlos completos, con sus resultados (inventados). Están aparte: <b>no se mezclan con tus indicadores</b>. Si uno te sirve, pulsa «Copiar a mi cuadro».</p>
      <div class="ej-grid">
        ${EJEMPLOS.map((e, i) => `<a class="ej-card" href="#/indicador/ej-${i + 1}/resultados"><b>${esc(e.titulo)}</b><span>${esc(e.sub)}</span></a>`).join("")}
        <a class="ej-card ej-card--cmp" href="#/compuesto/ej-c1"><b>Atención al infarto (compuesto)</b><span>Compuesto · 6 criterios · 70 %</span></a>
      </div>
      <div class="toolbar"><a class="btn btn--primary" href="#/ejemplos">Ver el cuadro de ejemplo completo</a></div>
      ${practica("Los ejemplos son ilustrativos: sus cifras y estándares son orientativos. Adáptalos siempre a tu unidad.")}`;
  }

  // ---------- Cuadro de indicadores ----------
  let filtroSub = "", orden = "ipr";
  function renderCuadro() {
    const informe = document.body.classList.contains("ver-informe");
    if (informe) return renderCuadroInforme();
    const subs = [...new Set(datos.indicadores.map(i => i.subproceso.trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
    let lista = datos.indicadores.filter(i => !filtroSub || i.subproceso.trim() === filtroSub);
    lista = ordenar(lista);
    const n = datos.indicadores.length;
    const diasCopia = datos.ultimaCopia ? Math.floor((Date.now() - new Date(datos.ultimaCopia)) / 864e5) : null;
    const avisoCopia = (n || datos.compuestos.length) && (diasCopia === null || diasCopia > 30)
      ? `<div class="warn">⚠ ${diasCopia === null ? "Todavía no has hecho ninguna copia de seguridad." : `Tu última copia de seguridad es de hace ${diasCopia} días.`} Tus datos solo están en este dispositivo. <button class="link-btn" data-act="copia">Hacer copia ahora</button></div>` : "";
    if (modoEjemplo) {
      view.innerHTML = `
      <h1 class="page-title">Cuadro de ejemplo</h1>
      <p class="page-lead">${n} indicadores y ${datos.compuestos.length} compuesto de muestra, con <b>resultados inventados</b> para que veas los semáforos, los gráficos y la tendencia. Puedes tocarlo todo: <b>no se guarda</b> y no se mezcla con tus indicadores.</p>
      <div class="toolbar">
        <a class="btn btn--primary" href="#/cuadro">← Ir a mi cuadro</a>
        <button class="btn" data-act="informe-cuadro">Ver / imprimir el cuadro</button><button class="btn" data-act="excel">Exportar a Excel</button>
      </div>
      ${equilibrioHtml()}
      <div class="cuadro-lista">${lista.map(tarjetaInd).join("")}</div>
      <h2 class="sec-t">Indicadores compuestos</h2><div class="cuadro-lista">${datos.compuestos.map(tarjetaCmp).join("")}</div>`;
      return;
    }
    const nomEj = new Set(EJEMPLOS.map(e => e.ind.nombre));
    const copiasEj = datos.indicadores.filter(i => nomEj.has(i.nombre)).length + datos.compuestos.filter(c => c.nombre === EJEMPLO_COMPUESTO.nombre).length;
    const avisoEj = copiasEj ? `<div class="note">ℹ Tu cuadro tiene <b>${copiasEj}</b> copia${copiasEj === 1 ? "" : "s"} de los ejemplos (añadidas con una versión anterior de la app). Ahora los ejemplos están aparte, en el <a href="#/ejemplos">cuadro de ejemplo</a>. <button class="link-btn" data-act="quitar-copias-ej">Quitarlas de mi cuadro</button></div>` : "";
    view.innerHTML = `
      <h1 class="page-title">Mi cuadro de indicadores</h1>
      <p class="page-lead">${n ? `${n} indicador${n === 1 ? "" : "es"}${datos.compuestos.length ? ` y ${datos.compuestos.length} compuesto${datos.compuestos.length === 1 ? "" : "s"}` : ""}. Se guardan solo en este dispositivo.` : "Todavía no tienes indicadores. Crea el primero o mira antes el cuadro de ejemplo."}</p>
      <div class="toolbar">
        <a class="btn btn--primary" href="#/nuevo">+ Nuevo indicador</a>
        <a class="btn" href="#/compuesto/nuevo">+ Nuevo compuesto</a>
        <a class="btn" href="#/ejemplos">Ver cuadro de ejemplo</a>
        ${n ? `<button class="btn" data-act="informe-cuadro">Ver / imprimir el cuadro</button><button class="btn" data-act="excel">Exportar a Excel</button>` : ""}
      </div>
      ${avisoEj}
      ${avisoCopia}
      ${n ? equilibrioHtml() : ""}
      ${n > 1 ? `<div class="filtros">
        <label>Subproceso <select data-filtro="sub"><option value="">Todos</option>${subs.map(s => `<option${s === filtroSub ? " selected" : ""}>${esc(s)}</option>`).join("")}</select></label>
        <label>Ordenar por <select data-filtro="orden">${[["ipr", "Prioridad (IPR)"], ["nombre", "Nombre"], ["reciente", "Lo último modificado"], ["tipo", "Tipo"]].map(([v, t]) => `<option value="${v}"${v === orden ? " selected" : ""}>${t}</option>`).join("")}</select></label>
      </div>` : ""}
      <div class="cuadro-lista">${lista.map(tarjetaInd).join("")}</div>
      ${datos.compuestos.length ? `<h2 class="sec-t">Indicadores compuestos</h2><div class="cuadro-lista">${datos.compuestos.map(tarjetaCmp).join("")}</div>` : ""}
      <div class="card copia">
        <h2 class="sub-h">Copia de seguridad</h2>
        <p class="field-sub">Guarda un archivo con todo tu cuadro para no perderlo si cambias de móvil o borras el navegador, o para pasarlo a otro dispositivo.</p>
        <div class="toolbar"><button class="btn" data-act="copia">⬇ Descargar copia (.json)</button><button class="btn" data-act="importar">⬆ Importar copia</button></div>
        ${datos.ultimaCopia ? `<p class="field-sub">Última copia: ${fechaCorta(datos.ultimaCopia)}.</p>` : ""}
      </div>`;
  }
  function ordenar(lista) {
    const ord = { estructura: 1, proceso: 2, resultado: 3, cumplimiento: 4 };
    const f = { ipr: (a, b) => (L.ipr(b.g, b.o, b.d) || 0) - (L.ipr(a.g, a.o, a.d) || 0), nombre: (a, b) => a.nombre.localeCompare(b.nombre, "es"), reciente: (a, b) => String(b.modificado).localeCompare(String(a.modificado)), tipo: (a, b) => (ord[a.tipo] || 9) - (ord[b.tipo] || 9) }[orden];
    return lista.slice().sort(f);
  }
  function equilibrioHtml() {
    const cuenta = {}; Object.keys(L.TIPOS).forEach(k => { cuenta[k] = 0; });
    let sin = 0; datos.indicadores.forEach(i => { if (cuenta[i.tipo] !== undefined) cuenta[i.tipo]++; else sin++; });
    const max = Math.max(1, ...Object.values(cuenta));
    const falta = ["estructura", "proceso", "resultado"].filter(k => !cuenta[k]);
    return `<div class="card equilibrio"><h2 class="sub-h">Equilibrio del cuadro</h2>
      <div class="eq-barras">${Object.keys(L.TIPOS).map(k => `<div class="eq-fila"><span class="eq-lbl">${L.TIPOS[k].nom}</span><span class="eq-track"><span class="eq-fill eq-fill--${k}" style="width:${cuenta[k] / max * 100}%"></span></span><b>${cuenta[k]}</b></div>`).join("")}</div>
      ${sin ? `<p class="field-sub">${sin} sin tipo asignado.</p>` : ""}
      ${falta.length && datos.indicadores.length >= 2 ? `<p class="field-sub">💡 Un cuadro equilibrado mira lo que se tiene, lo que se hace y lo que se logra. Te falta${falta.length > 1 ? "n" : ""}: <b>${falta.map(k => L.TIPOS[k].nom.toLowerCase()).join(", ")}</b>.</p>` : ""}
    </div>`;
  }
  function puntosProgreso(ind) {
    const c = L.pasosCompletos(ind);
    return `<span class="prog" title="Pasos completados">${[1, 2, 3, 4, 5].map(p => `<span class="prog-p${c[p] ? " ok" : ""}"></span>`).join("")}</span>`;
  }
  function tarjetaInd(ind) {
    const t = L.TIPOS[ind.tipo], c = L.CALCULOS[ind.calculo], ipr = L.ipr(ind.g, ind.o, ind.d), e = L.estadoCalidad(ind);
    const meta = [L.textoEstandar(ind), ind.periodicidad, ind.responsable].filter(lleno).map(esc).join(" · ");
    const ult = L.analizarResultados(ind).ultima;
    return `<div class="ind-card">
      <a class="ind-link" href="#/indicador/${ind.id}/${L.pasosCompletos(ind)[3] ? 6 : 1}">
        <div class="badges">${t ? `<span class="badge badge--${ind.tipo}">${t.letra}</span>` : ""}${c ? `<span class="badge">${c.nom}</span>` : ""}${ipr ? `<span class="badge badge--ipr">IPR ${ipr}</span>` : ""}<span class="badge badge--sem badge--${e.color}"><span class="punto"></span>${e.texto}</span></div>
        <b class="ind-nombre">${lleno(ind.nombre) ? esc(ind.nombre) : '<span class="vacio">Sin nombre</span>'}</b>
        ${lleno(ind.subproceso) || lleno(ind.hallazgo) ? `<span class="ind-sub">${esc([ind.subproceso, ind.hallazgo].filter(lleno).join(" · "))}</span>` : ""}
        ${meta ? `<span class="ind-meta">${meta}</span>` : ""}
        ${ult ? `<span class="res-est res-est--${ult.estado}"><span class="punto"></span>${esc(ult.periodo)}: ${esc(L.textoResultado(ind, ult))} · ${ESTADO_TXT(ind, ult)}</span>` : ""}
      </a>
      <div class="ind-pie">${puntosProgreso(ind)}<span class="ind-acc"><a href="#/indicador/${ind.id}/6">Ficha</a><a href="#/indicador/${ind.id}/resultados">Resultados</a>${modoEjemplo ? `<button class="link-btn" data-act="copiar-a-mio" data-id="${ind.id}">Copiar a mi cuadro</button>` : `<button class="link-btn" data-act="duplicar" data-id="${ind.id}">Duplicar</button><button class="link-btn link-btn--peligro" data-act="borrar" data-id="${ind.id}">Borrar</button>`}</span></div>
    </div>`;
  }
  function tarjetaCmp(c) {
    const r = L.calcCompuesto(c.rows);
    return `<div class="ind-card">
      <a class="ind-link" href="#/compuesto/${c.id}">
        <div class="badges"><span class="badge badge--cmp">Σ Compuesto</span>${r.valido ? `<span class="badge">${r.filas.length} criterios</span>` : ""}</div>
        <b class="ind-nombre">${lleno(c.nombre) ? esc(c.nombre) : '<span class="vacio">Compuesto sin nombre</span>'}</b>
        ${r.valido ? `<span class="ind-meta"><b class="cmp-mini">${L.pct1(r.pct)}</b> · ${r.cum}/${r.opo} oportunidades · punto débil: ${esc(r.flojo.nom)} (${L.pct1(r.flojo.pct)})</span>` : `<span class="ind-meta">Sin datos completos</span>`}
      </a>
      <div class="ind-pie"><span></span><span class="ind-acc">${modoEjemplo ? `<button class="link-btn" data-act="copiar-a-mio" data-id="${c.id}">Copiar a mi cuadro</button>` : `<button class="link-btn link-btn--peligro" data-act="borrar-cmp" data-id="${c.id}">Borrar</button>`}</span></div>
    </div>`;
  }
  function ultimoTxt(ind) {
    const u = L.analizarResultados(ind).ultima;
    return u ? `<span class="res-est res-est--${u.estado}"><span class="punto"></span>${ESTADO_TXT(ind, u)}</span><br>${esc(u.periodo)}: ${esc(L.textoResultado(ind, u))}` : "—";
  }
  function renderCuadroInforme() {
    const lista = ordenar(datos.indicadores);
    view.innerHTML = `${informeBar()}${informeCab("Cuadro de indicadores")}
      <div class="table-wrap"><table class="cuadro-t">
        <thead><tr><th>Nº</th><th>Indicador</th><th>Tipo · cálculo</th><th>Fórmula</th><th>Estándar</th><th>Fuente</th><th>Periodicidad · responsable</th><th>Plan</th><th>Último resultado</th></tr></thead>
        <tbody>${lista.map((i, k) => `<tr><td>${k + 1}</td><td><b>${esc(i.nombre)}</b>${lleno(i.hallazgo) ? `<br><span class="gris">${esc(i.subproceso)}${lleno(i.subproceso) ? " · " : ""}${esc(i.hallazgo)}${L.ipr(i.g, i.o, i.d) ? ` (IPR ${L.ipr(i.g, i.o, i.d)})` : ""}</span>` : ""}</td><td>${L.TIPOS[i.tipo] ? L.TIPOS[i.tipo].nom : "—"}<br><span class="gris">${L.CALCULOS[i.calculo] ? L.CALCULOS[i.calculo].nom : ""}</span></td><td>${esc(L.formula(i))}</td><td>${esc(L.textoEstandar(i))}</td><td>${esc(L.fuentesTexto(i))}</td><td>${esc(i.periodicidad)}<br><span class="gris">${esc(i.responsable)}</span></td><td>${esc(L.textoPlan(i))}</td><td>${ultimoTxt(i)}</td></tr>`).join("")}</tbody>
      </table></div>
      ${datos.compuestos.map(c => { const r = L.calcCompuesto(c.rows); return `<h2 class="sec-t">${esc(c.nombre || "Compuesto")}</h2>${r.valido ? desgloseTabla(r) + `<p class="an-frase-p">${esc(lleno(c.conclusion) ? c.conclusion : L.fraseCompuesto(c.nombre, r))}</p>` : "<p>Sin datos completos.</p>"}`; }).join("")}
      ${informePie()}`;
  }

  // ---------- Compuesto ----------
  let cmpActual = null;
  function renderCompuesto(c) {
    cmpActual = c; actual = null; nAyuda = 0;
    const informe = document.body.classList.contains("ver-informe");
    view.innerHTML = `
      ${informe ? informeBar() : ""}
      <div class="wiz-top no-print">${volverCuadro()}</div>
      <h1 class="page-title no-print">Indicador compuesto</h1>
      ${informeCab("Indicador compuesto" + (lleno(c.nombre) ? ": " + c.nombre : ""))}
      <div class="no-print">
      ${idea("Un compuesto agrupa varios criterios en un número: <b>Σ cumplimientos ÷ Σ oportunidades</b>. Es el titular; el desglose, la noticia. <b>Nunca lo presentes sin su desglose.</b>")}
      ${facil(`<p>No todos los criterios se aplican a todos los pacientes (el consejo antitabaco, solo en fumadores). Por eso cada criterio tiene su propio denominador: las <b>oportunidades</b> de actuar bien.</p>
        <ul><li><b>Pesan más los criterios que más se evalúan.</b> No siempre es malo, pero hay que saberlo.</li><li><b>Un buen titular puede esconder una mala noticia.</b> Mira siempre el criterio más flojo.</li><li><b>No todos los compuestos son iguales.</b> La media de los porcentajes da otro número: indica siempre el método.</li><li><b>Cuarenta oportunidades no son cuarenta pacientes.</b> Las de un mismo paciente se parecen.</li></ul>`)}
      </div>
      <div class="calc-grid">
        <div class="card no-print">
          ${campo(c, "nombre", "Nombre del compuesto", { ph: "p. ej. Paquete de valoración al ingreso" })}
          ${campo(c, "pacientes", "Pacientes revisados", { num: true, opt: true, unit: "pacientes", help: "Para recordar que las oportunidades no son pacientes." })}
          ${campo(c, "estandar", "Estándar del compuesto", { num: true, opt: true, unit: "%", help: "Si lo indicas, se marca en las barras." })}
          <div class="field"><div class="field-head"><label class="field-label">Criterios (subindicadores)</label></div>
            ${c.rows.map((r, i) => `<div class="cmp-row"><div class="cmp-row-head"><span class="cmp-num">${i + 1}</span><input class="cmp-nom txt-in" data-k="rows.${i}.nom" value="${esc(r.nom)}" placeholder="Criterio ${i + 1}" />${c.rows.length > 2 ? `<button class="icon-btn cmp-del" data-act="cmp-del" data-i="${i}" aria-label="Quitar criterio">✕</button>` : ""}</div>
              <div class="cmp-nums"><label>Cumplimientos<input class="txt-in" data-k="rows.${i}.cum" value="${esc(r.cum)}" inputmode="numeric" /></label><label>Oportunidades<input class="txt-in" data-k="rows.${i}.opo" value="${esc(r.opo)}" inputmode="numeric" /></label></div></div>`).join("")}
            <button class="btn cmp-add" data-act="cmp-add">+ Añadir criterio</button>
            ${datos.indicadores.some(i => i.calculo === "porcentaje") ? `<label class="cmp-desde">O añade uno de tu cuadro: <select data-act-sel="cmp-desde"><option value="">Elegir…</option>${datos.indicadores.filter(i => i.calculo === "porcentaje").map(i => `<option value="${i.id}">${esc(i.nombre.slice(0, 80))}</option>`).join("")}</select></label>` : ""}
          </div>
          <div class="toolbar">${esVacioCmp(c) ? `<a class="btn" href="#/compuesto/ej-c1">Ver el ejemplo (infarto)</a>` : ""}<a class="btn" href="${L.TM_URL}#/compuesto?d=${encodeURIComponent(JSON.stringify({ pac: c.pacientes, rows: c.rows.map(r => ({ nom: r.nom, cum: r.cum, opo: r.opo })) }))}" target="_blank" rel="noopener" data-zona-href-cmp>Abrir en Tamaño Muestral ↗</a></div>
        </div>
        <div class="calc-result-col" data-zona-cmp>${resultadoCmpHtml(c)}</div>
      </div>
      ${informePie()}`;
  }
  const esVacioCmp = c => !lleno(c.nombre) && c.rows.every(r => !lleno(r.nom) && !lleno(r.cum) && !lleno(r.opo));
  function desgloseTabla(r, estandar) {
    return `<div class="table-wrap"><table class="mini cmp-t"><thead><tr><th>Criterio</th><th>Cumplen</th><th>Oportunidades</th><th>%</th><th>Peso</th></tr></thead><tbody>
      ${r.filas.map(f => `<tr class="${f === r.flojo ? "flojo" : ""}"><td>${esc(f.nom)}</td><td>${f.cum}</td><td>${f.opo}</td><td>${L.pct1(f.pct)}</td><td>${L.pct1(f.peso)}</td></tr>`).join("")}
      <tr class="sel"><td>Indicador compuesto</td><td>${r.cum}</td><td>${r.opo}</td><td>${L.pct1(r.pct)}</td><td>100 %</td></tr></tbody></table></div>`;
  }
  function resultadoCmpHtml(c) {
    const r = L.calcCompuesto(c.rows), est = L.num(c.estandar);
    if (!r.valido) return `<div class="placeholder-box">${r.errores.map(esc).join("<br>")}</div>`;
    const frase = lleno(c.conclusion) ? c.conclusion : L.fraseCompuesto(c.nombre, r);
    const pac = L.num(c.pacientes);
    return `
      <div class="result-card"><div class="res-label">Compuesto (porcentaje simple)</div><div class="res-big">${L.pct1(r.pct)}</div>
        <div class="res-sub">${r.cum} cumplimientos de ${r.opo} oportunidades${pac ? ` en ${pac} pacientes` : ""}</div>
        <div class="extras"><span class="extra">Media de los % <b>${L.pct1(r.mediaPct)}</b></span><span class="extra">Punto débil <b>${L.pct1(r.flojo.pct)}</b></span></div>
      </div>
      <div class="chart-card"><h3>Desglose: el compuesto es el titular; esto, la noticia</h3>
        <div class="cmp-bars">${r.filas.map(f => `<div class="cmp-bar${f === r.flojo ? " cmp-bar--flojo" : ""}"><div class="cmp-bar-top"><span>${esc(f.nom)}</span><b>${L.pct1(f.pct)}</b></div><div class="cmp-track"><div class="cmp-fill" style="width:${f.pct}%"></div>${est !== null && !Number.isNaN(est) && est >= 0 && est <= 100 ? `<span class="cmp-mark" style="left:${est}%"></span>` : ""}</div><div class="cmp-bar-sub">${f.cum} de ${f.opo} · pesa el ${L.pct1(f.peso)} del compuesto</div></div>`).join("")}</div>
        ${est !== null && !Number.isNaN(est) ? `<p>La raya vertical marca el estándar (${nf(est)} %).</p>` : ""}
      </div>
      <div class="print-only">${desgloseTabla(r)}</div>
      ${Math.abs(r.mediaPct - r.pct) >= 0.05 ? `<div class="note">ℹ La media de los porcentajes (${L.pct1(r.mediaPct)}) no coincide con el compuesto (${L.pct1(r.pct)}): la media da el mismo peso a «${esc(r.filas.reduce((a, f) => f.opo < a.opo ? f : a, r.filas[0]).nom)}» (${r.filas.reduce((a, f) => f.opo < a.opo ? f : a, r.filas[0]).opo} oportunidades) que a «${esc(r.pesado.nom)}» (${r.pesado.opo}). El compuesto pondera por oportunidades.</div>` : ""}
      ${r.flojo.pct < r.pct - 10 ? `<div class="warn">⚠ Un buen titular puede esconder una mala noticia: «${esc(r.flojo.nom)}» está en el ${L.pct1(r.flojo.pct)}, muy por debajo del compuesto.</div>` : ""}
      ${pac && r.opo > pac ? `<div class="warn">⚠ ${r.opo} oportunidades no son ${r.opo} pacientes: vienen de ${pac} pacientes, y las de un mismo paciente se parecen.</div>` : ""}
      <div class="an-frase"><span class="box-tag">Conclusión (titular + desglose)</span><p>${esc(frase)}</p></div>
      <div class="no-print">${campo(c, "conclusion", "Tu conclusión", { area: true, opt: true, ph: L.fraseCompuesto(c.nombre, r), help: "Si la dejas vacía se usa la frase propuesta." })}</div>
      <div class="toolbar no-print"><button class="btn btn--primary" data-act="informe">Ver informe</button><button class="btn" data-act="imprimir">Guardar / imprimir</button>${modoEjemplo ? `<button class="btn" data-act="copiar-a-mio" data-id="${c.id}">Copiar a mi cuadro</button>` : `<button class="btn btn--peligro-ghost" data-act="borrar-cmp" data-id="${c.id}">Borrar</button>`}</div>`;
  }
  function refrescarCmp() {
    const z = $("[data-zona-cmp]"); if (!z || !cmpActual) return;
    // conserva el foco si se está escribiendo en «Tu conclusión»
    const foco = document.activeElement && document.activeElement.dataset && document.activeElement.dataset.k === "conclusion";
    if (!foco) z.innerHTML = resultadoCmpHtml(cmpActual);
    const a = $("[data-zona-href-cmp]"); if (a) a.href = `${L.TM_URL}#/compuesto?d=${encodeURIComponent(JSON.stringify({ pac: cmpActual.pacientes, rows: cmpActual.rows.map(r => ({ nom: r.nom, cum: r.cum, opo: r.opo })) }))}`;
  }

  // ---------- Compartir un indicador por enlace ----------
  function compacto(ind) {
    const o = {}, base = L.nuevoIndicador("", "");
    Object.keys(ind).forEach(k => {
      if (["id", "creado", "modificado"].includes(k)) return;
      const v = ind[k];
      if (typeof v === "string" ? v !== "" : Array.isArray(v) ? v.length : v && typeof v === "object" ? JSON.stringify(v) !== JSON.stringify(base[k]) && Object.keys(v).length : false) o[k] = v;
    });
    return o;
  }
  function enlaceIndicador(ind) { return location.href.split("#")[0] + "#/ver?d=" + encodeURIComponent(JSON.stringify(compacto(ind))); }
  async function compartirEnlace(ind) {
    const url = enlaceIndicador(ind), movil = esIOS() || /Android/i.test(navigator.userAgent);
    if (movil && navigator.share) { try { await navigator.share({ title: "Indicador · " + (ind.nombre || "sin nombre"), url }); } catch (e) { /* cancelado */ } return; }
    try { await navigator.clipboard.writeText(url); toast("Enlace copiado: quien lo abra verá la ficha y podrá añadirla a su cuadro.", 4500); }
    catch (e) { window.prompt("Copia este enlace:", url); }
  }
  let compartido = null;
  function renderVer() {
    if (!compartido) { view.innerHTML = `<div class="error-box">El enlace no contiene un indicador válido.</div><p><a href="#/">Ir al inicio</a></p>`; return; }
    view.innerHTML = `<div class="note">ℹ Te han compartido este indicador. Míralo y, si quieres, añádelo a tu cuadro para modificarlo.</div>
      <div class="toolbar"><button class="btn btn--primary" data-act="anadir-compartido">+ Añadir a mi cuadro</button></div>
      ${fichaHtml(compartido, { soloLectura: true })}`;
  }

  // ---------- Aprende ----------
  function renderAprende() {
    view.innerHTML = `
      <h1 class="page-title">Aprende</h1>
      <p class="page-lead">Lo esencial sobre indicadores, para tenerlo a mano mientras diseñas.</p>
      <div class="ciclo card">
        <h2 class="sub-h">El círculo completo</h2>
        <ol class="ciclo-l">${["Mapa de Riesgos (AMFE): qué puede fallar y con qué prioridad (IPR)", "Criterio: cómo debería hacerse bien", "Indicador: el criterio convertido en número", "Muestra: cuántos casos revisar (Tamaño Muestral)", "Auditoría: recoger las evidencias", "Informe: comparar con el estándar", "Mejora y seguimiento: actuar y volver a medir"].map(x => `<li>${x}</li>`).join("")}</ol>
        <p class="field-sub">Si un indicador detecta una desviación, se revisa también el Mapa de Riesgos: mapa, indicador y auditoría se retroalimentan.</p>
      </div>
      ${idea("Medir lo importante, medir bien y saber interpretar. Un indicador es un <b>termómetro</b>: muestra cómo está un proceso, avisa de desviaciones y sirve para decidir mejoras.")}
      <details class="facil" open><summary>1 · Dato o indicador</summary><div class="facil-body"><p>Un <b>dato</b> es una información puntual: «este mes se notificaron 12 incidencias». Un <b>indicador</b> relaciona la información con objetivos y estándares, permite ver tendencias y valorar si lo que se hace se ajusta a lo planificado.</p><p>Se construye siempre a partir de un <b>criterio</b>: «se registra la valoración del riesgo de úlceras por presión en las primeras 24 horas» → «% de pacientes ingresados con la valoración registrada en las primeras 24 horas».</p></div></details>
      <details class="facil"><summary>2 · Tipos según lo que miden</summary><div class="facil-body"><table class="mini facil-tabla"><tr><th>Tipo</th><th>Mira…</th><th>Ejemplo</th></tr>${Object.keys(L.TIPOS).map(k => `<tr><td><b>${L.TIPOS[k].nom}</b></td><td>${L.TIPOS[k].que.toLowerCase()}</td><td>${EJ_TIPO[k]}</td></tr>`).join("")}</table><p>Un cuadro <b>equilibrado</b> mira el proceso desde las cuatro perspectivas.</p></div></details>
      <details class="facil"><summary>3 · Tipos según su forma de cálculo</summary><div class="facil-body"><div class="table-wrap"><table class="mini facil-tabla"><tr><th>Tipo</th><th>Qué mide</th><th>Ejemplo</th><th>¿Muestra?</th></tr>${Object.keys(L.CALCULOS).map(k => `<tr><td><b>${L.CALCULOS[k].nom}</b></td><td>${L.CALCULOS[k].que}</td><td>${EJ_CALC[k]}</td><td>${L.CALCULOS[k].muestra}</td></tr>`).join("")}</table></div>
        ${trampa("La media puede maquillar: conviértela en «% de pacientes que superan el límite».")}
        <p>En auditoría de historias casi todo es <b>cumple / no cumple</b> (dicotómico): se expresa en porcentaje y se analiza con métodos sencillos y muestras pequeñas.</p></div></details>
      <details class="facil"><summary>4 · Un buen indicador: SMART, válido, fiable, útil</summary><div class="facil-body"><ul><li><b>Específico</b>: aspecto concreto y bien definido.</li><li><b>Medible</b>: cuantificable de forma objetiva.</li><li><b>Alcanzable</b>: viable con los recursos disponibles.</li><li><b>Relevante</b>: alineado con los objetivos clave.</li><li><b>Temporal</b>: periodicidad definida.</li></ul><p>Y tres preguntas más: <b>¿válido?</b> (mide lo que dice medir), <b>¿fiable?</b> (dos auditoras llegan a lo mismo) y <b>¿útil?</b> (alguien puede actuar). También debe ser <b>comparable</b> en el tiempo y entre unidades.</p>${trampa("Si un auditor marca «cumple» y otro «no cumple» en la misma historia, el problema está en el indicador. Pilótalo con dos personas antes de lanzarlo.")}</div></details>
      <details class="facil"><summary>5 · Seleccionar y definir</summary><div class="facil-body"><ol><li>Identificar áreas clave: Mapa de Riesgos, auditorías anteriores, objetivos de la unidad.</li><li>Elegir solo los que aporten información para decidir: ¿qué queremos saber? ¿qué decisión tomaremos?</li><li>Definir cada uno con precisión: nombre, fórmula, unidad, periodicidad, fuente, responsable y objetivo de referencia.</li><li>Verificar la viabilidad: ¿hay datos?, ¿se registran de forma fiable?, ¿hay quien los interprete?</li></ol><p>No se trata de listas interminables, sino de <b>priorizar</b>.</p></div></details>
      <details class="facil"><summary>6 · Indicadores compuestos</summary><div class="facil-body"><div class="formula">Compuesto (%) = Σ cumplimientos ÷ Σ oportunidades × 100</div>${ejemplo("infarto de miocardio", "<p>15 pacientes, 6 criterios, 40 oportunidades, 28 cumplidas → <b>70 %</b>. Pero el consejo antitabaco está en el 33 % y los β-bloqueantes al alta en el 50 %. La media de los seis porcentajes daría 63,9 %.</p>")}<p><b>Regla de oro:</b> nunca presentes el compuesto sin su desglose. <a href="#/compuesto/nuevo">Abrir el módulo de compuestos</a>.</p></div></details>
      <details class="facil"><summary>7 · Implantación, seguimiento y PDCA</summary><div class="facil-body"><ul><li><b>Planificar</b>: diseñar indicadores alineados con los riesgos.</li><li><b>Hacer</b>: recoger los datos (fuente y periodicidad claras).</li><li><b>Verificar</b>: analizar, comparar con el estándar, ver tendencias. En la app, desde la ficha: <b>Resultados</b> (un periodo por fila, semáforo y gráfico con la línea del estándar).</li><li><b>Actuar</b>: medidas de mejora y, si hace falta, ajustar el indicador.</li></ul><p>Los resultados se difunden de forma clara y orientada a la acción. Un indicador no es una carga burocrática: es una herramienta de mejora.</p></div></details>
      <details class="facil"><summary>8 · Estimar o decidir (y Tamaño Muestral)</summary><div class="facil-body"><p><b>Estimar</b> es la báscula de precisión: da un número con su margen de error, pero pide más historias (97 con ±10 %, 385 con ±5 %; menos si el universo es pequeño). <b>Decidir</b> es el arco del aeropuerto: el LQAS dice «hay o no hay problema» con 15–40 historias.</p><p>Esta app enlaza con <a href="${L.TM_URL}" target="_blank" rel="noopener">Tamaño Muestral ↗</a> para hacer esos cálculos con los datos de tu indicador.</p></div></details>`;
  }

  // ---------- Glosario ----------
  const GLOSARIO = [
    ["Indicador", "Herramienta que traduce la información de un proceso o resultado en un valor que se puede analizar, comparar con un objetivo y usar para decidir."],
    ["KPI", "Key Performance Indicator: indicador clave de rendimiento. Se centra en lo estratégico."],
    ["Dato", "Información puntual y aislada. Se convierte en indicador cuando se relaciona con un objetivo o estándar."],
    ["Criterio", "Regla de buena práctica que debería cumplirse. Es el punto de partida del indicador."],
    ["Mapa de Riesgos", "Herramienta que identifica los procesos críticos y lo que puede fallar en ellos. Orienta qué medir primero."],
    ["AMFE", "Análisis Modal de Fallos y Efectos: para cada fallo se valora gravedad, ocurrencia y detección."],
    ["IPR", "Índice de Prioridad del Riesgo = gravedad × ocurrencia × detección (cada una de 1 a 10). Cuanto más alto, más prioritario."],
    ["Hallazgo", "Fallo o problema detectado (en el Mapa de Riesgos o en una auditoría) que conviene vigilar."],
    ["Subproceso", "Parte del proceso asistencial donde se sitúa el riesgo (admisión, triaje, ingreso…)."],
    ["Indicador de estructura", "Mide lo que se tiene: personal, equipos, instalaciones, organización."],
    ["Indicador de proceso", "Mide lo que se hace: si se aplican los protocolos y procedimientos."],
    ["Indicador de resultado", "Mide lo que se logra: el efecto final sobre la salud y la seguridad del paciente."],
    ["Indicador de cumplimiento", "Mide si se respetan normas, estándares o acuerdos internos (por ejemplo, estándares ACSA)."],
    ["Centinela", "Indicador en el que un solo caso ya es un problema grave. No se muestrea: se vigilan todos los casos y cada uno se analiza."],
    ["Número absoluto", "Cuenta eventos sin denominador. No sirve para comparar unidades de distinto tamaño."],
    ["Porcentaje", "Casos que cumplen ÷ casos evaluados × 100. El más usado en auditoría de historias."],
    ["Tasa", "Casos referidos a una población y un periodo (por 1.000 estancias, por 100 ingresos…)."],
    ["Media", "Promedio de una medición numérica. Puede maquillar: conviene acompañarla o sustituirla por un porcentaje."],
    ["Ratio o razón", "Promedio de cumplimientos o defectos por caso (p. ej. campos sin rellenar por informe)."],
    ["Indicador compuesto", "Resume varios criterios en un número. El de porcentaje simple es Σ cumplimientos ÷ Σ oportunidades."],
    ["Oportunidad", "Cada ocasión en la que se podía haber actuado bien con un criterio. Es el denominador de cada criterio en un compuesto."],
    ["Desglose", "El porcentaje de cada criterio de un compuesto. El compuesto es el titular; el desglose, la noticia."],
    ["Numerador", "Lo que se cuenta arriba: los casos que cumplen o los eventos."],
    ["Denominador", "Sobre qué casos se calcula: los evaluados o la población de referencia."],
    ["Definición operativa", "Descripción precisa del indicador (fórmula, fuente, estándar, términos…) para que todos midan lo mismo."],
    ["Exclusiones", "Casos que no se cuentan, decididos de antemano."],
    ["Fuente de datos", "De dónde sale la información: historia clínica, registros, sistemas de información, notificaciones, observación…"],
    ["Periodicidad", "Cada cuánto se mide el indicador."],
    ["Estándar", "Nivel que nos comprometemos a alcanzar (meta). A veces se llama «umbral» o «límite aceptable»."],
    ["Umbral de alarma", "En LQAS, nivel por debajo del cual la situación es claramente mala."],
    ["Sentido del indicador", "Si es mejor que suba (cumplimientos: estándar con ≥) o que baje (fallos: estándar con ≤)."],
    ["Benchmark", "Valor de referencia de otras unidades o centros con el que compararse."],
    ["SMART", "Específico, Medible, Alcanzable, Relevante y Temporal."],
    ["Validez", "Que el indicador mida de verdad lo que pretende medir."],
    ["Fiabilidad", "Que mida siempre lo mismo en las mismas condiciones: dos auditoras con la misma historia llegan a lo mismo."],
    ["Utilidad", "Que quien recibe el resultado pueda hacer algo para mejorarlo."],
    ["Comparabilidad", "Que se pueda analizar en el tiempo y entre unidades parecidas, con el mismo método."],
    ["Pilotaje", "Probar el indicador con unas pocas historias revisadas por dos personas antes de lanzarlo."],
    ["Kappa", "Medida de acuerdo entre dos auditoras que descuenta el acuerdo debido al azar. Se calcula en Tamaño Muestral."],
    ["Dicotómico", "Que solo tiene dos respuestas: cumple / no cumple."],
    ["Muestra aleatoria", "Casos elegidos de forma que todos tenían la misma probabilidad de salir. La de conveniencia no vale."],
    ["LQAS", "Muestreo para la aceptación de lotes: con pocas historias decide si se acepta o se rechaza un estándar. No estima el cumplimiento."],
    ["Cuadro de indicadores", "Conjunto de indicadores de una unidad o proceso, equilibrado entre tipos."],
    ["PDCA", "Ciclo de mejora: Planificar, Hacer, Verificar y Actuar. Los indicadores alimentan el Verificar."]
  ];
  function renderGlosario() {
    view.innerHTML = `<h1 class="page-title">Glosario</h1><p class="page-lead">${GLOSARIO.length} términos en lenguaje llano.</p>
      <div class="input-wrap glos-buscar"><input id="glos-q" type="search" placeholder="Buscar término…" autocomplete="off" /></div>
      <dl class="glos-lista" id="glos-lista"></dl>`;
    const pinta = q => {
      const t = L.sinTildes(q);
      const l = GLOSARIO.filter(([a, b]) => !t || L.sinTildes(a + " " + b).includes(t)).sort((x, y) => x[0].localeCompare(y[0], "es"));
      $("#glos-lista").innerHTML = l.length ? l.map(([a, b]) => `<div class="glos-item"><dt>${esc(a)}</dt><dd>${esc(b)}</dd></div>`).join("") : `<p class="placeholder-box">Sin resultados.</p>`;
    };
    pinta(""); $("#glos-q").addEventListener("input", e => pinta(e.target.value));
  }

  // ---------- Ponte a prueba ----------
  const QUIZ = [
    { q: "«Cirugía en el lado equivocado»: ¿se mide con una muestra de historias?", o: ["Sí, con LQAS", "Sí, con 50 historias al azar", "No: es centinela, se vigilan todos los casos"], ok: 2, e: "Un solo caso ya es un problema grave: se vigilan todos y cada uno se analiza." },
    { q: "«Ratio de enfermeras por número de camas» es un indicador de…", o: ["Estructura", "Proceso", "Resultado"], ok: 0, e: "Mide un recurso disponible: lo que se tiene." },
    { q: "«% de pacientes a los que se aplica la lista de verificación quirúrgica» es de…", o: ["Estructura", "Proceso", "Resultado"], ok: 1, e: "Mide si se hace una actividad del protocolo." },
    { q: "«Tasa de infección de localización quirúrgica» es de…", o: ["Proceso", "Resultado", "Cumplimiento"], ok: 1, e: "Mide el efecto final sobre el paciente." },
    { q: "Dos auditoras revisan la misma historia: una marca «cumple» y la otra «no cumple». ¿Dónde está el problema?", o: ["En la historia", "En una de las auditoras", "En la definición del indicador"], ok: 2, e: "Hay que aclarar los términos ambiguos y pilotar el indicador antes de lanzarlo." },
    { q: "El tiempo medio de espera es de 20 minutos. ¿Qué indicador es mejor?", o: ["El mismo: la media lo resume todo", "% de pacientes que esperan más de 30 minutos", "El tiempo máximo del mes"], ok: 1, e: "La media puede maquillar pacientes de tres horas; el porcentaje dice cuántos quedan fuera del límite." },
    { q: "«% de errores de identificación en las peticiones de laboratorio». ¿Qué estándar tiene sentido?", o: ["≥ 95 %", "≤ 2 %", "= 50 %"], ok: 1, e: "Cuenta fallos: menos es mejor y el estándar va con ≤." },
    { q: "«Mejorar la seguridad del paciente en la planta». ¿Qué letra de SMART falla primero?", o: ["Específico", "Relevante", "Ninguna"], ok: 0, e: "Es una intención, no un aspecto concreto: no se puede medir así." },
    { q: "El resultado de un indicador no lo puede cambiar nadie del equipo. Le falta ser…", o: ["Válido", "Útil", "Medible"], ok: 1, e: "Un indicador que nadie puede modificar solo genera trabajo." },
    { q: "«Nº de caídas con daño al mes»: ¿sirve para comparar una planta de 20 camas con otra de 40?", o: ["Sí, sin problema", "No: mejor una tasa por 1.000 estancias", "Solo si se mide cada semana"], ok: 1, e: "Sin denominador, la planta grande parecerá peor aunque lo haga igual." },
    { q: "Compuesto de porcentaje simple: ¿cómo se calcula?", o: ["Media de los porcentajes de cada criterio", "Σ cumplimientos ÷ Σ oportunidades × 100", "% de pacientes con todo correcto"], ok: 1, e: "Los otros métodos dan otros números: indica siempre el método usado." },
    { q: "El compuesto sale 89 %. ¿Qué presentas a la Comisión?", o: ["El 89 %: es lo que importa", "El 89 % con su desglose por criterios", "Solo el criterio peor"], ok: 1, e: "El compuesto es el titular; el desglose, la noticia." }
  ];
  function renderPrueba() {
    view.innerHTML = `<h1 class="page-title">Ponte a prueba</h1><p class="page-lead">Toca la respuesta que creas correcta. <span id="quiz-marcador"></span></p>
      <div class="card">${QUIZ.map((q, i) => `<div class="quiz-q"><b>${i + 1}.</b> ${esc(q.q)}${q.o.map((o, j) => `<button class="quiz-o" data-quiz="${i}" data-o="${j}">${esc(o)}</button>`).join("")}<p class="quiz-fb" hidden></p></div>`).join("")}</div>`;
  }
  const quizAciertos = {};

  // ---------- PDF en iOS (sin librerías; mismo método que Tamaño Muestral) ----------
  // En iOS, window.print() no hace nada cuando la app está instalada en la pantalla de inicio.
  function esIOS() { return /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1); }
  function toast(msg, ms) {
    let t = $("#toast");
    if (!msg) { if (t) t.hidden = true; return; }
    if (!t) { t = document.createElement("div"); t.id = "toast"; t.className = "toast"; t.setAttribute("role", "status"); document.body.appendChild(t); }
    t.textContent = msg; t.hidden = false; clearTimeout(toast.h);
    if (ms) toast.h = setTimeout(() => { t.hidden = true; }, ms);
  }
  function pdfDesdeJPEGs(paginas, wPt, hPt) {
    const enc = new TextEncoder(), partes = [], offs = []; let pos = 0;
    const push = x => { const b = typeof x === "string" ? enc.encode(x) : x; partes.push(b); pos += b.length; };
    const obj = (n, cuerpo) => { offs[n] = pos; push(`${n} 0 obj\n`); cuerpo(); push("\nendobj\n"); };
    push("%PDF-1.4\n"); push(new Uint8Array([37, 226, 227, 207, 211, 10]));
    const n = paginas.length;
    obj(1, () => push("<< /Type /Catalog /Pages 2 0 R >>"));
    obj(2, () => push(`<< /Type /Pages /Kids [${paginas.map((_, i) => `${3 + i * 3} 0 R`).join(" ")}] /Count ${n} >>`));
    paginas.forEach((p, i) => {
      const pg = 3 + i * 3, cont = pg + 1, img = pg + 2, flujo = `q ${wPt} 0 0 ${hPt} 0 0 cm /Im${i} Do Q`;
      obj(pg, () => push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${wPt} ${hPt}] /Resources << /XObject << /Im${i} ${img} 0 R >> >> /Contents ${cont} 0 R >>`));
      obj(cont, () => { push(`<< /Length ${flujo.length} >>\nstream\n${flujo}\nendstream`); });
      obj(img, () => { push(`<< /Type /XObject /Subtype /Image /Width ${p.w} /Height ${p.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${p.bytes.length} >>\nstream\n`); push(p.bytes); push("\nendstream"); });
    });
    const total = 3 + n * 3, xref = pos;
    let x = `xref\n0 ${total}\n0000000000 65535 f \n`;
    for (let k = 1; k < total; k++) x += String(offs[k]).padStart(10, "0") + " 00000 n \n";
    push(x + `trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    return new Blob(partes, { type: "application/pdf" });
  }
  async function generarPDF() {
    const ANCHO = 720, K = 2;
    const A4W = 595.28, A4H = 841.89, MARGEN_PT = 34;
    const PX = (A4W - 2 * MARGEN_PT) / ANCHO;
    const PW = A4W / PX, PH = A4H / PX, M = MARGEN_PT / PX, ALTO_UTIL = PH - 2 * M;
    const hoja = document.createElement("div");
    hoja.className = "view";
    hoja.style.cssText = `position:absolute;left:-20000px;top:0;width:${ANCHO}px;max-width:none;margin:0;padding:0;box-shadow:none;`;
    hoja.innerHTML = $("#view").innerHTML;
    $$(".informe-bar, [id]", hoja).forEach(el => { if (el.classList.contains("informe-bar")) el.remove(); else el.removeAttribute("id"); });
    $$("details", hoja).forEach(d => { d.open = true; });
    document.body.appendChild(hoja);
    try {
      await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
      const R0 = hoja.getBoundingClientRect();
      const ops = [], cortes = [0], bloques = [];
      const SIN_CORTE = ".print-head, .result-card, .warn, .note, .chart-card, tr, .formula, .box, .print-foot, h1, h2, h3, .ficha-cab, .an-frase, .cmp-bar";
      const rel = r => ({ x: r.left - R0.left, y: r.top - R0.top, w: r.width, h: r.height });
      const hayColor = c => !!c && c !== "transparent" && !/rgba\([^)]*,\s*0\)$/.test(c);
      const LADOS = ["Top", "Right", "Bottom", "Left"];
      const texto = (n, cs) => {
        const t = n.textContent, re = /\S+/g, rg = document.createRange(); let m;
        const font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
        while ((m = re.exec(t))) {
          rg.setStart(n, m.index); rg.setEnd(n, m.index + m[0].length);
          const rs = rg.getClientRects(); if (!rs.length) continue;
          ops.push({ t: "txt", s: cs.textTransform === "uppercase" ? m[0].toUpperCase() : m[0], r: rel(rs[0]), font, col: cs.color, ls: cs.letterSpacing });
        }
      };
      const recorre = el => {
        const cs = getComputedStyle(el);
        if (cs.display === "none" || cs.visibility === "hidden") return;
        const r = rel(el.getBoundingClientRect());
        if (r.h > 0) cortes.push(r.y + r.h);
        if (el.matches(SIN_CORTE)) bloques.push([r.y, r.y + r.h]);
        if (el instanceof SVGSVGElement) { ops.push({ t: "svg", el, r }); return; }
        const bordes = LADOS.map(l => [parseFloat(cs[`border${l}Width`]) || 0, cs[`border${l}Color`]]);
        if (hayColor(cs.backgroundColor) || bordes.some(b => b[0] > 0 && hayColor(b[1]))) ops.push({ t: "caja", r, bg: cs.backgroundColor, rad: parseFloat(cs.borderTopLeftRadius) || 0, bordes });
        if (el.tagName === "LI") {
          const ol = el.parentElement.tagName === "OL", idx = Array.prototype.indexOf.call(el.parentElement.children, el) + 1;
          ops.push({ t: "txt", s: ol ? idx + "." : "•", r: { x: r.x - (ol ? 20 : 14), y: r.y, w: 16, h: parseFloat(cs.lineHeight) || parseFloat(cs.fontSize) * 1.4 }, font: `normal 700 ${cs.fontSize} ${cs.fontFamily}`, col: ol ? "#0b6f66" : cs.color, ls: "normal" });
        }
        for (const n of el.childNodes) {
          if (n.nodeType === 1) recorre(n);
          else if (n.nodeType === 3 && n.textContent.trim()) texto(n, cs);
        }
      };
      recorre(hoja);
      for (const op of ops.filter(o => o.t === "svg")) {
        const cl = op.el.cloneNode(true), orig = [op.el, ...op.el.querySelectorAll("*")], copia = [cl, ...cl.querySelectorAll("*")];
        orig.forEach((o, i) => { const cs = getComputedStyle(o); copia[i].setAttribute("style", ["fill", "stroke", "stroke-width", "stroke-dasharray", "font-family", "font-size", "font-weight"].map(p => `${p}:${cs.getPropertyValue(p)}`).join(";")); });
        cl.setAttribute("xmlns", "http://www.w3.org/2000/svg"); cl.setAttribute("width", op.r.w); cl.setAttribute("height", op.r.h);
        op.img = await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.onerror = () => res(null); im.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(new XMLSerializer().serializeToString(cl)); });
      }
      const alto = Math.ceil(R0.height);
      const parte = y => bloques.some(([a, b]) => a < y - 0.5 && y + 0.5 < b && b - a < ALTO_UTIL);
      const cands = [...new Set(cortes.map(c => Math.ceil(c)))].sort((a, b) => a - b);
      const pags = []; let y0 = 0;
      while (y0 < alto - 1) {
        const lim = y0 + ALTO_UTIL;
        if (lim >= alto) { pags.push([y0, alto]); break; }
        const y = cands.filter(c => c > y0 + 80 && c <= lim && !parte(c)).pop() || Math.floor(lim);
        pags.push([y0, y]); y0 = y;
      }
      const ponFuente = (ctx, f) => { ctx.font = "10px serif"; ctx.font = f; if (ctx.font === "10px serif") ctx.font = f.replace(/^(.*?\d+(\.\d+)?px).*$/, "$1 sans-serif"); };
      const caja = (ctx, op) => {
        const { x, y, w, h } = op.r, rad = Math.min(op.rad, w / 2, h / 2);
        const camino = () => { ctx.beginPath(); if (rad > 0 && ctx.roundRect) ctx.roundRect(x, y, w, h, rad); else ctx.rect(x, y, w, h); };
        if (hayColor(op.bg)) { ctx.fillStyle = op.bg; camino(); ctx.fill(); }
        const [t, d, b, l] = op.bordes;
        if (op.bordes.every(z => z[0] === t[0] && z[1] === t[1]) && t[0] > 0 && hayColor(t[1])) { ctx.strokeStyle = t[1]; ctx.lineWidth = t[0]; ctx.beginPath(); if (rad > 0 && ctx.roundRect) ctx.roundRect(x + t[0] / 2, y + t[0] / 2, w - t[0], h - t[0], Math.max(0, rad - t[0] / 2)); else ctx.rect(x + t[0] / 2, y + t[0] / 2, w - t[0], h - t[0]); ctx.stroke(); return; }
        const lado = (z, rx, ry, rw, rh) => { if (z[0] > 0 && hayColor(z[1])) { ctx.fillStyle = z[1]; ctx.fillRect(rx, ry, rw, rh); } };
        lado(t, x, y, w, t[0]); lado(b, x, y + h - b[0], w, b[0]); lado(l, x, y, l[0], h); lado(d, x + w - d[0], y, d[0], h);
      };
      const dibujaPaginas = conSVG => pags.map(([a, b]) => {
        const cv = document.createElement("canvas"); cv.width = Math.round(PW * K); cv.height = Math.round(PH * K);
        const ctx = cv.getContext("2d");
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);
        ctx.scale(K, K); ctx.translate(M, M - a);
        ctx.beginPath(); ctx.rect(-2, a, ANCHO + 4, b - a); ctx.clip();
        for (const op of ops) {
          if (op.r.y >= b || op.r.y + op.r.h <= a) continue;
          if (op.t === "caja") caja(ctx, op);
          else if (op.t === "svg") { if (conSVG && op.img) ctx.drawImage(op.img, op.r.x, op.r.y, op.r.w, op.r.h); }
          else {
            ponFuente(ctx, op.font); ctx.fillStyle = op.col; ctx.textBaseline = "alphabetic";
            if ("letterSpacing" in ctx) ctx.letterSpacing = op.ls && op.ls !== "normal" ? op.ls : "0px";
            const mt = ctx.measureText(op.s), asc = mt.fontBoundingBoxAscent || mt.actualBoundingBoxAscent || op.r.h * 0.8, desc = mt.fontBoundingBoxDescent || mt.actualBoundingBoxDescent || op.r.h * 0.2;
            ctx.fillText(op.s, op.r.x, op.r.y + (op.r.h - (asc + desc)) / 2 + asc);
          }
        }
        const url = cv.toDataURL("image/jpeg", 0.9), bin = atob(url.slice(url.indexOf(",") + 1)), bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return { bytes, w: cv.width, h: cv.height };
      });
      let paginas;
      try { paginas = dibujaPaginas(true); } catch (e) { paginas = dibujaPaginas(false); }
      return pdfDesdeJPEGs(paginas, A4W, A4H);
    } finally { hoja.remove(); }
  }
  let pdfListo = null;
  const rutaActual = () => location.hash.split("?")[0];
  function nombreArchivo(ext) {
    const base = actual ? "Ficha " + (actual.nombre || "indicador").slice(0, 50) : cmpActual ? "Compuesto " + (cmpActual.nombre || "").slice(0, 50) : "Cuadro de indicadores";
    return `${base.replace(/[\\/:*?"<>|%]/g, "").trim()} ${new Date().toLocaleDateString("es-ES").replace(/\//g, "-")}.${ext}`;
  }
  function prepararPDF() {
    pdfListo = null; const r = rutaActual();
    setTimeout(() => { if (rutaActual() === r && document.body.classList.contains("ver-informe")) generarPDF().then(b => { pdfListo = { r, blob: b }; }).catch(() => {}); }, 300);
  }
  async function guardarPDF() {
    if (!document.body.classList.contains("ver-informe")) { irInforme(); toast("Preparando el PDF… pulsa de nuevo «Guardar / imprimir» en un momento.", 4000); return; }
    let blob = pdfListo && pdfListo.r === rutaActual() ? pdfListo.blob : null;
    if (!blob) { toast("Preparando el PDF…"); try { blob = await generarPDF(); pdfListo = { r: rutaActual(), blob }; } catch (e) { toast("No se ha podido crear el PDF.", 4000); return; } }
    const file = new File([blob], nombreArchivo("pdf"), { type: "application/pdf" });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: file.name }); toast(""); }
      catch (e) { if (e && e.name === "AbortError") { toast(""); return; } toast("PDF listo: pulsa otra vez «Guardar / imprimir».", 4000); }
      return;
    }
    descargar(blob, file.name); toast("PDF descargado.", 3000);
  }
  function descargar(blob, nombre) {
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 200);
  }
  function irInforme() {
    if (actual) location.hash = `#/indicador/${actual.id}/6?informe=1`;
    else if (cmpActual) location.hash = `#/compuesto/${cmpActual.id}?informe=1`;
    else location.hash = rutaCuadro() + "?informe=1";
  }
  function imprimir() {
    if (esIOS()) { guardarPDF(); return; }
    if (!document.body.classList.contains("ver-informe")) { irInforme(); setTimeout(() => window.print(), 400); return; }
    window.print();
  }

  // ---------- Copia de seguridad, importar y Excel ----------
  async function copiaSeguridad() {
    const contenido = JSON.stringify({ app: "indicadores", v: L.VERSION_DATOS, exportado: hoy(), indicadores: datosUsuario.indicadores, compuestos: datosUsuario.compuestos }, null, 1);
    const nombre = `Indicadores copia ${new Date().toISOString().slice(0, 10)}.json`;
    const blob = new Blob([contenido], { type: "application/json" });
    const file = new File([blob], nombre, { type: "application/json" });
    if (esIOS() && navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: nombre }); } catch (e) { if (e && e.name === "AbortError") return; }
    } else descargar(blob, nombre);
    datosUsuario.ultimaCopia = hoy(); guardar(true);
    toast("Copia guardada. Guárdala en un sitio seguro (Archivos, Drive, correo…).", 4500);
    if (rutaActual() === rutaCuadro()) renderCuadro();
  }
  function importar(file) {
    const fr = new FileReader();
    fr.onload = () => {
      let s = null; try { s = L.sanearDatos(JSON.parse(fr.result), genId, hoy()); } catch (e) { /* */ }
      if (!s) { toast("Ese archivo no es una copia de Indicadores.", 4000); return; }
      const hay = datos.indicadores.length + datos.compuestos.length;
      const aplicar = modo => {
        const r = modo === "sustituir" ? s : L.fusionar(datosUsuario, s);
        datosUsuario.indicadores = r.indicadores; datosUsuario.compuestos = r.compuestos; guardar(true);
        toast(`Importados ${s.indicadores.length} indicadores y ${s.compuestos.length} compuestos.`, 4000);
        location.hash = "#/cuadro"; route();
      };
      if (!hay) { aplicar("sustituir"); return; }
      confirmar(`<p>La copia tiene <b>${s.indicadores.length}</b> indicadores y <b>${s.compuestos.length}</b> compuestos.</p><p><b>Añadir</b> los junta con los tuyos (si uno está en los dos, se queda el más reciente). Si prefieres <b>sustituir</b> todo tu cuadro por la copia, pulsa «Sustituir».</p>`, "Añadir a mi cuadro", () => aplicar("anadir"), { titulo: "Importar copia", peligro: false, otro: ["Sustituir", () => confirmar("<p>Se borrará tu cuadro actual y se quedará solo lo de la copia. Esto no se puede deshacer.</p>", "Sustituir", () => aplicar("sustituir"))] });
    };
    fr.readAsText(file);
  }
  function exportarExcel() {
    try {
      const libro = L.hojasExcel({ indicadores: ordenar(datos.indicadores), compuestos: datos.compuestos }, new Date().toLocaleDateString("es-ES"));
      const bytes = window.Xlsx.construirLibroXlsx(libro);
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const nombre = `Cuadro de indicadores ${new Date().toISOString().slice(0, 10)}.xlsx`;
      const file = new File([blob], nombre, { type: blob.type });
      if (esIOS() && navigator.canShare && navigator.canShare({ files: [file] })) navigator.share({ files: [file], title: nombre }).catch(() => {});
      else { descargar(blob, nombre); toast("Excel descargado.", 3000); }
    } catch (e) { toast("No se ha podido crear el Excel.", 4000); }
  }

  // ---------- Confirmación ----------
  function confirmar(html, siTxt, alSi, o = {}) {
    const m = $("#modal-confirmar");
    $("#conf-titulo").textContent = o.titulo || "¿Seguro?";
    $("#conf-texto").innerHTML = html;
    const si = $("#btn-conf-si"), no = $("#btn-conf-no");
    si.textContent = siTxt; si.className = "btn " + (o.peligro === false ? "btn--primary" : "btn--peligro");
    let extra = $("#btn-conf-otro"); if (extra) extra.remove();
    if (o.otro) { extra = document.createElement("button"); extra.id = "btn-conf-otro"; extra.className = "btn"; extra.textContent = o.otro[0]; no.after(extra); extra.onclick = () => { m.hidden = true; o.otro[1](); }; }
    si.onclick = () => { m.hidden = true; alSi(); };
    no.onclick = () => { m.hidden = true; };
    m.hidden = false; si.focus();
  }

  // ---------- Enrutado ----------
  const NAV = [["", "Inicio"], ["cuadro", "Mi cuadro"], ["ejemplos", "Ejemplos"], ["nuevo", "Nuevo indicador"], ["compuesto/nuevo", "Compuesto"], ["aprende", "Aprende"], ["glosario", "Glosario"], ["prueba", "Ponte a prueba"]];
  function buildNav() { $("#calc-nav").innerHTML = NAV.map(([r, t]) => `<a href="#/${r}" data-r="${r}">${t}</a>`).join(""); }
  function route() {
    const [ruta, qs] = (location.hash.replace(/^#\/?/, "") || "").split("?");
    const partes = ruta.split("/");
    const informe = /(^|&)informe=1(&|$)/.test(qs || "");
    document.body.classList.toggle("ver-informe", informe);
    actual = null; cmpActual = null;
    const idR = partes[1] || "";
    modoEjemplo = partes[0] === "ejemplos" || (["indicador", "compuesto"].includes(partes[0]) && esIdEjemplo(idR) && !!(cuadroEjemplo().indicadores.some(x => x.id === idR) || cuadroEjemplo().compuestos.some(x => x.id === idR)));
    datos = modoEjemplo ? cuadroEjemplo() : datosUsuario;
    const banner = $("#banner-ejemplo");
    banner.hidden = !modoEjemplo || informe || partes[0] === "ejemplos";
    $("#banner-copiar").hidden = partes[0] === "ejemplos";
    let navR = partes[0];
    if (partes[0] === "nuevo") {
      const ind = L.nuevoIndicador(genId(), hoy());
      borradores.set(ind.id, ind);
      history.replaceState(null, "", `#/indicador/${ind.id}/1`);
      renderWizard(ind, 1); navR = "nuevo";
    } else if (partes[0] === "indicador") {
      const ind = buscaInd(partes[1]) || borradores.get(partes[1]);
      const paso = Math.min(6, Math.max(1, parseInt(partes[2], 10) || 1));
      if (ind && partes[2] === "resultados" && !informe) renderResultados(ind);
      else if (ind) renderWizard(ind, informe ? 6 : paso); else { view.innerHTML = `<div class="error-box">No se encuentra ese indicador en este dispositivo.</div><p><a href="#/cuadro">Ir a mi cuadro</a></p>`; }
      navR = buscaInd(partes[1]) ? "cuadro" : "nuevo";
    } else if (partes[0] === "compuesto") {
      let c = partes[1] === "nuevo" ? null : buscaCmp(partes[1]) || borradores.get(partes[1]);
      if (!c) { c = L.nuevoCompuesto(genId(), hoy()); borradores.set(c.id, c); history.replaceState(null, "", `#/compuesto/${c.id}`); }
      renderCompuesto(c); navR = "compuesto/nuevo";
    } else if (partes[0] === "cuadro" || partes[0] === "ejemplos") renderCuadro();
    else if (partes[0] === "ver") {
      const md = /(?:^|&)d=([^&]*)/.exec(qs || "");
      compartido = null;
      if (md) { try { compartido = L.sanearIndicador(JSON.parse(decodeURIComponent(md[1])), genId(), hoy()); } catch (e) { compartido = null; } }
      history.replaceState(null, "", "#/ver");
      if (!md && compartidoPrevio) compartido = compartidoPrevio;
      compartidoPrevio = compartido;
      renderVer();
    }
    else if (partes[0] === "aprende") renderAprende();
    else if (partes[0] === "glosario") renderGlosario();
    else if (partes[0] === "prueba") renderPrueba();
    else { renderHome(); navR = ""; }
    if (modoEjemplo) navR = "ejemplos";
    if (informe) { $$("details").forEach(d => { d.open = true; }); if (esIOS()) prepararPDF(); }
    $$("#calc-nav a").forEach(a => a.classList.toggle("active", a.dataset.r === navR));
    const act = $("#calc-nav a.active"); if (act && act.scrollIntoView) act.scrollIntoView({ block: "nearest", inline: "center" });
    window.scrollTo(0, 0);
  }
  const borradores = new Map();
  let compartidoPrevio = null;

  function copiarAMiCuadro(id) {
    const ej = cuadroEjemplo(), ind = ej.indicadores.find(x => x.id === id), cmp = ej.compuestos.find(x => x.id === id);
    if (ind) {
      const copia = Object.assign(clon(ind), { id: genId(), creado: hoy(), modificado: hoy(), resultados: [] });
      datosUsuario.indicadores.push(copia); guardar(true); location.hash = `#/indicador/${copia.id}/6`;
      toast("Copiado a tu cuadro (sin los resultados inventados). Ya es tuyo: cámbialo a tu gusto.", 4500);
    } else if (cmp) {
      const copia = Object.assign(clon(cmp), { id: genId(), creado: hoy(), modificado: hoy() });
      datosUsuario.compuestos.push(copia); guardar(true); location.hash = `#/compuesto/${copia.id}`;
      toast("Compuesto copiado a tu cuadro.", 3500);
    }
  }

  // ---------- Eventos ----------
  function tocado(obj) {
    obj.modificado = hoy();
    if (actual === obj && !esVacio(obj)) asegurarGuardado(obj, datos.indicadores);
    if (cmpActual === obj && !esVacioCmp(obj)) asegurarGuardado(obj, datos.compuestos);
    guardar();
  }
  // Los cuadros de texto crecen con lo escrito
  function autoAlto(el) { el.style.height = "auto"; el.style.height = el.scrollHeight + 2 + "px"; }
  new MutationObserver(() => $$("textarea.txt-in", view).forEach(autoAlto)).observe(view, { childList: true, subtree: true });
  view.addEventListener("input", e => {
    if (e.target.matches("textarea.txt-in")) autoAlto(e.target);
    const el = e.target.closest("[data-k]"); if (!el) return;
    const obj = actual || cmpActual; if (!obj) return;
    setPath(obj, el.dataset.k, el.value);
    $$(`[data-set="${el.dataset.k}"]`).forEach(b => b.classList.toggle("active", b.dataset.v === el.value));
    tocado(obj);
    if (actual) refrescarZonas(); else refrescarCmp();
  });
  view.addEventListener("change", e => {
    const f = e.target.closest("[data-filtro]");
    if (f) { if (f.dataset.filtro === "sub") filtroSub = f.value; else orden = f.value; renderCuadro(); return; }
    const s = e.target.closest("[data-act-sel='cmp-desde']");
    if (s && s.value && cmpActual) {
      const ind = buscaInd(s.value);
      if (ind) { const vacia = cmpActual.rows.find(r => !lleno(r.nom) && !lleno(r.cum) && !lleno(r.opo)); if (vacia) vacia.nom = ind.nombre; else cmpActual.rows.push({ nom: ind.nombre, cum: "", opo: "", ind: ind.id }); tocado(cmpActual); renderCompuesto(cmpActual); }
    }
  });
  view.addEventListener("toggle", e => {
    if (e.target.matches && e.target.matches("details.facil") && !document.body.classList.contains("ver-informe") && rutaActual().startsWith("#/indicador")) escribir(KEY_FACIL, e.target.open ? "1" : "0");
  }, true);
  view.addEventListener("click", e => {
    const t = e.target; let b;
    if ((b = t.closest(".info-btn"))) { const h = $("#help-" + b.dataset.info); if (!h) return; const open = h.hidden; h.hidden = !open; b.setAttribute("aria-expanded", String(open)); return; }
    if ((b = t.closest(".quiz-o"))) {
      const i = Number(b.dataset.quiz), q = QUIZ[i], caja = b.closest(".quiz-q"), j = Number(b.dataset.o), fb = $(".quiz-fb", caja);
      $$(".quiz-o", caja).forEach(x => x.classList.remove("quiz-o--ok", "quiz-o--mal"));
      b.classList.add(j === q.ok ? "quiz-o--ok" : "quiz-o--mal");
      fb.innerHTML = (j === q.ok ? "<b>¡Correcto!</b> " : "<b>No.</b> ") + esc(q.e); fb.hidden = false;
      if (!(i in quizAciertos)) quizAciertos[i] = j === q.ok;
      const n = Object.keys(quizAciertos).length, ok = Object.values(quizAciertos).filter(Boolean).length;
      $("#quiz-marcador").textContent = `Aciertos a la primera: ${ok} de ${n}${n === QUIZ.length ? ok >= 10 ? " · ¡Muy bien!" : " · Repasa «Aprende»" : ""}`;
      return;
    }
    const obj = actual || cmpActual;
    if ((b = t.closest("[data-set]")) && obj) { setPath(obj, b.dataset.set, b.dataset.v); tocado(obj); if (actual) renderWizard(actual, pasoActual); else renderCompuesto(cmpActual); return; }
    if ((b = t.closest("[data-pick]")) && obj) {
      const path = b.dataset.pick, v = b.dataset.v;
      setPath(obj, path, getPath(obj, path) === v && path !== "objetivo" && !path.startsWith("calidad") ? "" : v);
      if (path === "calculo" && v === "centinela") { obj.sentido = "menos"; obj.estandar = "0"; obj.objetivo = "sin"; }
      if (path === "calculo" && v !== "centinela" && obj.estandar === "0" && obj.objetivo === "sin" && obj.sentido === "menos") { obj.estandar = ""; obj.objetivo = ""; }
      tocado(obj);
      const y = window.scrollY; renderWizard(actual, pasoActual); window.scrollTo(0, y);
      return;
    }
    if ((b = t.closest("[data-multi]")) && actual) {
      const arr = actual[b.dataset.multi] = actual[b.dataset.multi] || [], v = b.dataset.v, i = arr.indexOf(v);
      if (i >= 0) arr.splice(i, 1); else arr.push(v);
      b.classList.toggle("active", i < 0); tocado(actual); refrescarZonas(); return;
    }
    if (!(b = t.closest("[data-act]"))) return;
    const a = b.dataset.act;
    if (a === "sugerir-nombre") {
      if (!lleno(actual.criterio)) { toast("Escribe primero el criterio.", 3000); return; }
      const hacer = () => { actual.nombre = L.sugerirNombre(actual.criterio); tocado(actual); renderWizard(actual, 1); toast("Revisa el nombre: cambia «casos» por lo que cuentas (pacientes, historias…).", 4500); };
      if (lleno(actual.nombre)) confirmar("<p>Se sustituirá el nombre que has escrito.</p>", "Sustituir", hacer, { peligro: false }); else hacer();
    }
    else if (a === "ind-lc" && actual) {
      const ind = actual;
      traerLC("Traer de Lectura Crítica", x => {
        const r = L.aplicarLectura(ind, x); tocado(ind); renderWizard(ind, 1);
        toast(r.criterio ? "Criterio escrito desde la pregunta PICO y evidencia guardada en «Observaciones» (paso 3). Revisa el criterio." : "Evidencia guardada en «Observaciones» (paso 3). El criterio no se ha cambiado.", 5000);
      });
    }
    else if (a === "copiar-a-mio" || (a === "duplicar" && modoEjemplo)) copiarAMiCuadro(b.dataset.id || (actual ? actual.id : cmpActual ? cmpActual.id : ""));
    else if (a === "quitar-copias-ej") {
      const nomEj = new Set(EJEMPLOS.map(e => e.ind.nombre));
      const ni = datosUsuario.indicadores.filter(i => nomEj.has(i.nombre)).length, nc = datosUsuario.compuestos.filter(c => c.nombre === EJEMPLO_COMPUESTO.nombre).length;
      confirmar(`<p>Se quitarán de tu cuadro <b>${ni + nc}</b> copias de los ejemplos (las que conservan el mismo nombre que el ejemplo). Si has modificado alguna y quieres conservarla, cámbiale antes el nombre.</p><p>Los ejemplos seguirán disponibles en el cuadro de ejemplo.</p>`, "Quitarlas", () => {
        datosUsuario.indicadores = datosUsuario.indicadores.filter(i => !nomEj.has(i.nombre));
        datosUsuario.compuestos = datosUsuario.compuestos.filter(c => c.nombre !== EJEMPLO_COMPUESTO.nombre);
        guardar(true); renderCuadro(); toast("Copias de los ejemplos quitadas de tu cuadro.", 3000);
      }, { titulo: "Quitar copias de ejemplos" });
    }
    else if (a === "informe") irInforme();
    else if (a === "informe-cuadro") location.hash = rutaCuadro() + "?informe=1";
    else if (a === "volver") location.hash = rutaActual();
    else if (a === "imprimir") imprimir();
    else if (a === "enlace") compartirEnlace(actual);
    else if (a === "duplicar") {
      const orig = b.dataset.id ? buscaInd(b.dataset.id) : actual; if (!orig) return;
      const copia = Object.assign(clon(orig), { id: genId(), creado: hoy(), modificado: hoy(), nombre: (orig.nombre || "") + " (copia)" });
      datos.indicadores.push(copia); guardar(); location.hash = `#/indicador/${copia.id}/1`; toast("Copia creada: cámbiale lo que necesites.", 3000);
    }
    else if (a === "borrar") {
      const ind = b.dataset.id ? buscaInd(b.dataset.id) : actual; if (!ind) return;
      confirmar(`<p>Se borrará «${esc(ind.nombre || "indicador sin nombre")}» de este dispositivo. No se puede deshacer (salvo que tengas una copia de seguridad).</p>`, "Borrar", () => {
        datos.indicadores = datos.indicadores.filter(x => x.id !== ind.id); borradores.delete(ind.id); guardar(true);
        if (rutaActual() === rutaCuadro()) renderCuadro(); else location.hash = rutaCuadro();
        toast("Indicador borrado.", 2500);
      }, { titulo: "Borrar indicador" });
    }
    else if (a === "borrar-cmp") {
      const c = buscaCmp(b.dataset.id) || cmpActual; if (!c) return;
      confirmar(`<p>Se borrará el compuesto «${esc(c.nombre || "sin nombre")}». No se puede deshacer.</p>`, "Borrar", () => {
        datos.compuestos = datos.compuestos.filter(x => x.id !== c.id); borradores.delete(c.id); guardar(true);
        if (rutaActual() === rutaCuadro()) renderCuadro(); else location.hash = rutaCuadro();
      }, { titulo: "Borrar compuesto" });
    }
    else if (a === "cmp-add") { cmpActual.rows.push({ nom: "", cum: "", opo: "", ind: "" }); tocado(cmpActual); renderCompuesto(cmpActual); const ins = $$(".cmp-nom"); ins[ins.length - 1].focus(); }
    else if (a === "cmp-del") { cmpActual.rows.splice(Number(b.dataset.i), 1); tocado(cmpActual); renderCompuesto(cmpActual); }
    else if (a === "res-add" && actual) {
      const prev = actual.resultados[actual.resultados.length - 1];
      actual.resultados.push(L.nuevaMedicion(prev ? L.siguientePeriodo(prev.periodo) : ""));
      tocado(actual); renderResultados(actual);
      const ins = $$(".res-per"); const ultimo = ins[ins.length - 1];
      if (ultimo) { ultimo.scrollIntoView({ block: "center" }); (ultimo.value ? $(".res-nums input", ultimo.closest(".res-row")) || ultimo : ultimo).focus(); }
    }
    else if (a === "res-del" && actual) {
      const i = Number(b.dataset.i), m = actual.resultados[i]; if (!m) return;
      const quitar = () => { actual.resultados.splice(i, 1); tocado(actual); const y = window.scrollY; renderResultados(actual); window.scrollTo(0, y); };
      if ([m.num, m.den, m.valor, m.nota].some(lleno)) confirmar(`<p>Se quitará el periodo «${esc(m.periodo || "sin nombre")}» con sus datos.</p>`, "Quitar", quitar, { titulo: "Quitar periodo" }); else quitar();
    }
    else if (a === "copia") copiaSeguridad();
    else if (a === "importar") $("#input-importar").click();
    else if (a === "excel") exportarExcel();
    else if (a === "anadir-compartido" && compartido) {
      const ind = clon(compartido);
      if (esIdEjemplo(ind.id) || datosUsuario.indicadores.some(x => x.id === ind.id)) ind.id = genId();
      ind.creado = ind.modificado = hoy();
      datosUsuario.indicadores.push(ind); guardar(true); location.hash = `#/indicador/${ind.id}/6`; toast("Añadido a tu cuadro.", 3000);
    }
  });
  $("#banner-copiar").addEventListener("click", () => copiarAMiCuadro(actual ? actual.id : cmpActual ? cmpActual.id : ""));
  $("#input-importar").addEventListener("change", e => { const f = e.target.files && e.target.files[0]; if (f) importar(f); e.target.value = ""; });

  // Imprimir: abrir los desplegables y restaurarlos después
  let cerrados = [];
  window.addEventListener("beforeprint", () => { cerrados = $$("details").filter(d => !d.open); cerrados.forEach(d => { d.open = true; }); });
  window.addEventListener("afterprint", () => { cerrados.forEach(d => { d.open = false; }); cerrados = []; });

  // ---------- «Traer de Lectura Crítica» (mismo código en AMFE, DAFO, Auditorías e Indicadores) ----------
  // traerLC(título de la ventana, función que recibe la lectura elegida). Lee el paquete de este navegador
  // o, si no hay, ofrece cargar la copia .json de Lectura Crítica (#input-lc).
  let lcCargadas = null, lcPendiente = null;
  function traerLC(titulo, alElegir) {
    lcPendiente = { titulo, alElegir };
    // [eco]
    let r = null; try { r = L.lecturasDe(JSON.parse(leer("lecturacritica-intercambio") || "null")); } catch (e) { /* */ }
    const lista = lcCargadas || (r && r.lecturas && r.lecturas.length ? r.lecturas : null);
    if (lista) { elegirLC(lista); return; }
    confirmar(r && r.error === "version" ? "<p>Las lecturas de este navegador son de una versión de <b>Lectura Crítica</b> que esta app aún no sabe leer.</p>"
      : "<p>No hay lecturas de la app <b>Lectura Crítica</b> en este navegador.</p><p>Si las tienes en otro dispositivo, carga su copia de seguridad (.json).</p>",
      "Cargar una copia", () => $("#input-lc").click(), { titulo, peligro: false });
    // [/eco]
    /* [ind]
    if (lcCargadas) { elegirLC(lcCargadas); return; }
    $("#input-lc").click();
    [/ind] */
  }
  function elegirLC(lista) {
    const { titulo, alElegir } = lcPendiente, cortar = (s, n) => s.length > n ? s.slice(0, n - 1) + "…" : s;
    const fila = (x, i) => `<label class="lc-op"><input type="radio" name="lc-sel" value="${i}"${i ? "" : " checked"} /><span><b>${esc(cortar(x.titulo || L.referenciaLC(x), 110))}</b>`
      + `<small>${esc([x.tipoNombre, x.anio].filter(lleno).join(" · "))}${x.tipoNombre || x.anio ? " · " : ""}${x.decision ? `<span class="lc-dec lc-dec--${x.decision}">${L.DECISION_LC[x.decision]}</span>` : "Sin decisión"}</small></span></label>`;
    confirmar(`<p>Elige la lectura cuya evidencia quieres traer. Se copia como texto, que luego puedes cambiar.</p><div class="lc-lista">${lista.map(fila).join("")}</div>`, "Traer", () => {
      const s = document.querySelector('input[name="lc-sel"]:checked'), x = s && lista[Number(s.value)]; if (!x) return;
      if (x.decision === "aplicar") { alElegir(x); return; }
      confirmar(`<p>La decisión de esta lectura fue <b>${x.decision ? L.DECISION_LC[x.decision].toLowerCase() : "ninguna todavía"}</b>: con ella sola no basta para cambiar la práctica.</p><p>¿Traerla de todos modos?</p>`,
        "Traer igualmente", () => alElegir(x), { titulo, peligro: false });
    }, { titulo, peligro: false });
  }
  $("#input-lc").addEventListener("change", e => {
    const f = e.target.files && e.target.files[0]; e.target.value = ""; if (!f) return;
    const fr = new FileReader();
    fr.onload = () => {
      let r = null; try { r = L.lecturasDe(JSON.parse(fr.result)); } catch (er) { /* */ }
      if (!r || r.error) { toast(r && r.error === "version" ? "Esa copia es de una versión de Lectura Crítica que esta app aún no sabe leer." : "Ese archivo no es una copia de la app Lectura Crítica.", 4500); return; }
      if (!r.lecturas.length) { toast("Esa copia no tiene lecturas con título o referencia.", 4000); return; }
      lcCargadas = r.lecturas; if (lcPendiente) elegirLC(lcCargadas);
    };
    fr.readAsText(f);
  });

  // ---------- Tema ----------
  function setTema(t) {
    document.documentElement.setAttribute("data-tema", t);
    const m = $("#meta-theme-color"); if (m) m.content = t === "claro" ? "#f0f7fb" : "#0d1b2a";
  }
  setTema(document.documentElement.getAttribute("data-tema") || "claro");
  $("#btn-tema").addEventListener("click", () => {
    const t = document.documentElement.getAttribute("data-tema") === "oscuro" ? "claro" : "oscuro";
    setTema(t); escribir("indicadores-tema", t);
  });
  if (window.matchMedia) matchMedia("(prefers-color-scheme: dark)").addEventListener("change", ev => { if (!leer("indicadores-tema")) setTema(ev.matches ? "oscuro" : "claro"); });

  // ---------- Modales ----------
  const KEY_BIENV = "indicadores-bienvenida-vista";
  const mb = $("#modal-bienvenida"), ma = $("#modal-acerca");
  const vista = leer(KEY_BIENV);
  if (!vista) mb.hidden = false;
  $("#btn-bienv-ok").addEventListener("click", () => { if ($("#modal-no-repetir").checked) escribir(KEY_BIENV, "1"); mb.hidden = true; });
  $("#btn-acerca").addEventListener("click", () => { ma.hidden = false; });
  $("#btn-acerca-cerrar").addEventListener("click", () => { ma.hidden = true; });
  ma.addEventListener("click", e => { if (e.target === ma) ma.hidden = true; });
  document.addEventListener("keydown", e => { if (e.key === "Escape") { ma.hidden = true; $("#modal-confirmar").hidden = true; } });

  // ---------- Versiones y novedades ----------
  const KEY_VERSION = "indicadores-version-vista";
  const listaCambios = n => `<div class="version-item"><div class="version-cab"><b>${esc(n.titulo)}</b><span class="version-tag">v${esc(n.version)} · ${esc(n.fecha)}</span></div><ul>${n.cambios.map(c => `<li>${esc(c)}</li>`).join("")}</ul></div>`;
  $("#historial-versiones").innerHTML = NOVEDADES.map(listaCambios).join("");
  $("#about-version").textContent = $("#footer-version").textContent = `Versión ${APP_VERSION}`;
  const mn = $("#modal-novedades");
  $("#btn-novedades-ok").addEventListener("click", () => { mn.hidden = true; escribir(KEY_VERSION, APP_VERSION); });
  (function gestionarNovedades() {
    const vistaV = leer(KEY_VERSION);
    if (vistaV === APP_VERSION) return;
    if (!vistaV && !vista) { escribir(KEY_VERSION, APP_VERSION); return; }
    const idx = NOVEDADES.findIndex(n => n.version === vistaV);
    const nuevas = idx > 0 ? NOVEDADES.slice(0, idx) : [NOVEDADES[0]];
    $("#novedades-sub").textContent = nuevas.length > 1 ? "Esto es lo nuevo desde tu última visita" : "Gracias por mantener la app al día";
    $("#novedades-body").innerHTML = nuevas.map(listaCambios).join("");
    if (mb.hidden) mn.hidden = false;
  })();

  // ---------- Service worker y aviso de nueva versión ----------
  let actualizando = false;
  function avisoActualizacion(reg) {
    if ($("#update-banner")) return;
    const b = document.createElement("div");
    b.id = "update-banner"; b.className = "update-banner no-print"; b.setAttribute("role", "status");
    b.innerHTML = "<span>Nueva versión disponible</span><button type=\"button\">Actualizar</button>";
    b.querySelector("button").addEventListener("click", () => { guardar(true); actualizando = true; if (reg.waiting) reg.waiting.postMessage({ type: "SKIP_WAITING" }); b.remove(); });
    document.body.appendChild(b);
  }
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    navigator.serviceWorker.register("sw.js").then(reg => {
      if (reg.waiting && navigator.serviceWorker.controller) avisoActualizacion(reg);
      reg.addEventListener("updatefound", () => {
        const nw = reg.installing; if (!nw) return;
        nw.addEventListener("statechange", () => { if (nw.state === "installed" && navigator.serviceWorker.controller) avisoActualizacion(reg); });
      });
    }).catch(() => {});
    let recargando = false;
    navigator.serviceWorker.addEventListener("controllerchange", () => { if (recargando || !actualizando) return; recargando = true; location.reload(); });
  }

  // ---------- Instalar como aplicación ----------
  let avisoInstalar = null;
  window.addEventListener("beforeinstallprompt", e => { e.preventDefault(); avisoInstalar = e; const b = $("#btn-instalar"); if (b) b.hidden = false; });
  $("#btn-instalar").addEventListener("click", async () => { if (!avisoInstalar) return; avisoInstalar.prompt(); try { await avisoInstalar.userChoice; } catch (e) { /* */ } avisoInstalar = null; $("#btn-instalar").hidden = true; });
  window.addEventListener("appinstalled", () => { const b = $("#btn-instalar"); if (b) b.hidden = true; });
  if (matchMedia("(display-mode: standalone)").matches || navigator.standalone) { const i = $("#instalar-bloque"); if (i) i.classList.add("instalada"); }

  window.Indicadores = { generarPDF, datos: () => datos };   // solo para pruebas desde la consola

  // ---------- Arranque ----------
  buildNav();
  window.addEventListener("hashchange", route);
  route();
})();
