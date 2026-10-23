// ============================================================
//  logic.js — Lógica PURA (sin DOM) · Indicadores
//  Reglas del curso: tipos, formas de cálculo, fórmula, avisos y trampas,
//  calidad (SMART + válido, fiable, útil, comparable), indicador compuesto,
//  enlaces a Tamaño Muestral, validación de lo importado y hojas de Excel.
//  window.Logica en navegador y module.exports en Node (tests).
// ============================================================
(function (raiz) {
  "use strict";

  const TM_URL = "https://cjgaland.github.io/tamano-muestral/";
  const AUD_URL = "https://cjgaland.github.io/auditorias/";
  const AMFE_URL = "https://cjgaland.github.io/amfe/";
  const FLUJO_URL = "https://cjgaland.github.io/flujograma/";
  const DAFO_URL = "https://cjgaland.github.io/dafo/";
  const MAPA_URL = "https://cjgaland.github.io/mapa-riesgos/";
  const VERSION_DATOS = 1;

  // ---------- Catálogos ----------
  // Según lo que miden (tema 3, 3.1)
  const TIPOS = {
    estructura: { nom: "Estructura", letra: "E", que: "Lo que se tiene", def: "Recursos y medios: personal, equipos, instalaciones, organización." },
    proceso: { nom: "Proceso", letra: "P", que: "Lo que se hace", def: "Cómo se llevan a cabo las actividades: si se aplican los protocolos." },
    resultado: { nom: "Resultado", letra: "R", que: "Lo que se logra", def: "El efecto final sobre la salud y la seguridad del paciente." },
    cumplimiento: { nom: "Cumplimiento", letra: "C", que: "Cómo se respeta lo acordado", def: "Si se siguen normas, estándares o acuerdos internos." }
  };
  // Según su forma de cálculo (tema 3, tabla 1)
  const CALCULOS = {
    centinela: { nom: "Centinela", muestra: "no", que: "Un solo caso ya es un problema grave" },
    absoluto: { nom: "Número absoluto", muestra: "normalmente no", que: "Cuántos eventos ocurren, sin denominador" },
    porcentaje: { nom: "Porcentaje o proporción", muestra: "sí", que: "Casos que cumplen entre los casos evaluados" },
    tasa: { nom: "Tasa", muestra: "sí", que: "Casos referidos a una población y un periodo de tiempo" },
    media: { nom: "Media", muestra: "sí", que: "Promedio de una medición numérica" },
    ratio: { nom: "Ratio o razón", muestra: "sí", que: "Promedio de cumplimientos o defectos por caso" },
    compuesto: { nom: "Compuesto", muestra: "sí", que: "Resume varios indicadores en un solo número" }
  };
  const CALIDAD = [
    { k: "E", nom: "Específico", grupo: "SMART" },
    { k: "M", nom: "Medible", grupo: "SMART" },
    { k: "A", nom: "Alcanzable", grupo: "SMART" },
    { k: "R", nom: "Relevante", grupo: "SMART" },
    { k: "T", nom: "Temporal", grupo: "SMART" },
    { k: "valido", nom: "Válido", grupo: "Además" },
    { k: "fiable", nom: "Fiable", grupo: "Además" },
    { k: "util", nom: "Útil", grupo: "Además" },
    { k: "comparable", nom: "Comparable", grupo: "Además" }
  ];
  const PERIODICIDADES = ["Continua (cada caso)", "Diaria", "Semanal", "Mensual", "Trimestral", "Semestral", "Anual"];
  const FUENTES = ["Historia clínica", "Registro específico", "Sistema de información", "Notificación de incidentes", "Observación directa", "Encuesta", "Entrevista"];

  // ---------- Utilidades ----------
  function num(s) {
    if (s === undefined || s === null) return null;
    const t = String(s).trim().replace(",", ".");
    if (t === "") return null;
    return /^[-+]?(\d+\.?\d*|\.\d+)$/.test(t) ? Number(t) : NaN;
  }
  const lleno = s => typeof s === "string" ? s.trim() !== "" : s !== undefined && s !== null;
  const nfmt = x => Number.isFinite(x) ? String(Math.round(x * 100) / 100).replace(".", ",") : "—";
  const sinTildes = s => String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");

  // ---------- Modelo ----------
  function nuevoIndicador(id, fecha) {
    return {
      id, creado: fecha, modificado: fecha,
      // Paso 1 · Del riesgo al indicador
      subproceso: "", hallazgo: "", g: "", o: "", d: "", criterio: "", nombre: "",
      // Paso 2 · Caracterizar (5A)
      tipo: "", justTipo: "", calculo: "",
      // Paso 3 · Definición operativa (5B)
      numerador: "", denominador: "", base: "", unidad: "", poblacion: "", exclusiones: "",
      fuentes: [], fuenteOtra: "", sentido: "", estandar: "", umbral: "", periodicidad: "", responsable: "",
      terminos: "", observaciones: "",
      // Paso 4 · Calidad
      calidad: {}, piloto: { a: "", b: "", c: "", d: "" },
      // Paso 5 · Muestra
      objetivo: "", muestra: { conf: "95", e: "10", p: "", N: "", loss: "", alpha: "5", beta: "20", sigma: "", eMedia: "" },
      plan: { n: "", c: "", nota: "" },
      // Fase 2 · Resultados por periodo (Verificar del PDCA)
      resultados: []
    };
  }
  function nuevoCompuesto(id, fecha) {
    return { id, creado: fecha, modificado: fecha, nombre: "", pacientes: "", estandar: "", rows: [{ nom: "", cum: "", opo: "", ind: "" }, { nom: "", cum: "", opo: "", ind: "" }], conclusion: "" };
  }

  // ---------- Paso 1 ----------
  // IPR del AMFE del Mapa de Riesgos: gravedad × ocurrencia × detección (cada una de 1 a 10)
  function ipr(g, o, d) {
    const v = [g, o, d].map(num);
    if (v.some(x => x === null || !Number.isInteger(x) || x < 1 || x > 10)) return null;
    return v[0] * v[1] * v[2];
  }
  // Propuesta de nombre a partir del criterio («se registra X en las primeras 24 h» → «% de casos en los que se registra…»)
  function sugerirNombre(criterio) {
    let c = String(criterio || "").trim().replace(/[.;]+$/, "");
    if (!c) return "";
    c = c.charAt(0).toLowerCase() + c.slice(1);
    return "% de casos en los que " + c;
  }

  // ¿El texto describe algo malo (errores, eventos, demoras…)? Entonces «menos es mejor».
  const NEGATIVOS = /(error|discrepanc|incidente|incidencia|evento adverso|efecto adverso|caida|infeccion|ulcera|upp\b|reingreso|superan|supera |demora|retraso|omision|fallo|perdida|mortalidad|muerte|complicac|fuera de rango|mas de \d|superior a|sin (registr|valorac|reevaluac|conciliac|identific|consentim|pulsera|firma|alerta)|no consta|no se|incumpl|reclamac|flebitis|extravasac|lugar equivocado|paciente equivocado|daño|dano)/;
  function sentidoSugerido(texto) {
    // «riesgo de caídas», «prevención de UPP», «escala de…» no son eventos: se quitan antes de buscar
    const t = sinTildes(texto).replace(/riesgo (de |del |de la |de las |de los )?[a-z]+/g, " ").replace(/(riesgo|prevencion|escala|protocolo|valoracion|cribado|deteccion|notificacion|profilaxis) (de |del |de la )?[a-z]+( [a-z]+)?/g, " ");
    if (!sinTildes(texto).trim()) return null;
    return NEGATIVOS.test(t) ? "menos" : "mas";
  }

  // ---------- Paso 3 ----------
  function unidadPorDefecto(ind) {
    switch (ind.calculo) {
      case "porcentaje": case "compuesto": return "%";
      case "tasa": return "por " + (lleno(ind.base) ? ind.base : "1.000") + " " + (lleno(ind.denominador) ? ind.denominador : "unidades de tiempo-población");
      case "centinela": case "absoluto": return "casos";
      case "ratio": return "por caso";
      default: return "";
    }
  }
  function unidad(ind) { return lleno(ind.unidad) ? ind.unidad.trim() : unidadPorDefecto(ind); }

  function formula(ind) {
    const N = lleno(ind.numerador) ? ind.numerador.trim() : "[numerador]";
    const D = lleno(ind.denominador) ? ind.denominador.trim() : "[denominador]";
    const per = lleno(ind.periodicidad) ? " (" + ind.periodicidad.toLowerCase() + ")" : "";
    switch (ind.calculo) {
      case "porcentaje": return `(Nº de ${N} ÷ Nº de ${D}) × 100`;
      case "tasa": return `(Nº de ${N} ÷ ${D}) × ${lleno(ind.base) ? ind.base : "1.000"}`;
      case "media": return `Σ ${N} ÷ Nº de ${D}`;
      case "ratio": return `Nº de ${N} ÷ Nº de ${D}`;
      case "absoluto": return `Nº de ${N} en el periodo${per}`;
      case "centinela": return `Nº de casos de ${N}: cada caso se notifica y se analiza`;
      case "compuesto": return "(Σ cumplimientos ÷ Σ oportunidades) × 100";
      default: return "";
    }
  }
  const usaDenominador = c => ["porcentaje", "tasa", "media", "ratio"].includes(c);

  // Estándar con su signo: «≥ 95 %», «≤ 5 %», «0 casos»
  function textoEstandar(ind, campo = "estandar") {
    const v = num(ind[campo]);
    if (ind.calculo === "centinela" && campo === "estandar") return "0 casos (cualquier caso es un aviso)";
    if (v === null || Number.isNaN(v)) return "";
    const signo = v === 0 && ind.sentido === "menos" ? "" : ind.sentido === "menos" ? "≤" : ind.sentido === "mas" ? "≥" : "";
    const u = unidad(ind);
    return `${signo ? signo + " " : ""}${nfmt(v)}${u === "%" ? " %" : u ? " " + u : ""}`;
  }

  // ---------- Avisos ----------
  // Devuelve [{ nivel: "trampa" | "nota" | "error", paso, texto }]
  function avisos(ind) {
    const A = [], add = (nivel, paso, texto) => A.push({ nivel, paso, texto });
    const c = ind.calculo, est = num(ind.estandar), umb = num(ind.umbral);
    const sug = sentidoSugerido([ind.nombre, ind.numerador].join(" "));

    // Paso 2: según la forma de cálculo
    if (c === "centinela") {
      add("trampa", 2, "Un indicador centinela no se muestrea: se vigilan todos los casos. Un solo caso ya obliga a actuar (notificar y analizar las causas).");
      if (ind.tipo === "estructura") add("nota", 2, "Los centinela suelen ser de resultado (un evento grave). Revisa si encaja como estructura.");
    }
    if (c === "absoluto") add("nota", 2, "Un número absoluto no tiene denominador: si cambia el volumen de actividad, el número cambia aunque la calidad sea la misma. No sirve para comparar unidades de distinto tamaño; si puedes, conviértelo en tasa.");
    if (c === "media") add("trampa", 2, "La media puede maquillar: 20 minutos de espera media pueden esconder pacientes que esperan tres horas. Plantéate convertirla en porcentaje: «% de pacientes que esperan más de X minutos».");
    if (c === "tasa") add("nota", 2, "Una tasa necesita una población y un periodo: indica la base (por 100, por 1.000…) y en qué se mide (estancias, ingresos, pacientes-día…).");
    if (c === "compuesto") add("nota", 2, "Para el compuesto hay un módulo propio: en el cuadro pulsa «Nuevo compuesto». Aquí puedes definir su ficha.");
    if (c === "porcentaje") add("nota", 2, "Es dicotómico (cumple / no cumple): el más usado en auditoría de historias. Se analiza con métodos sencillos y muestras pequeñas.");

    // Paso 3: coherencia del sentido y del estándar
    if (ind.sentido && sug && sug !== ind.sentido) {
      if (sug === "menos") add("trampa", 3, "El nombre parece medir algo malo (errores, eventos, demoras…), pero has marcado «más es mejor». Si mide fallos, cuanto más bajo, mejor: el estándar va con ≤.");
      else add("nota", 3, "Has marcado «menos es mejor». Comprueba que el indicador mide fallos o eventos y no cumplimientos.");
    }
    if (est !== null && !Number.isNaN(est) && ["porcentaje", "compuesto"].includes(c)) {
      if (est < 0 || est > 100) add("error", 3, "Un estándar en porcentaje debe estar entre 0 y 100.");
      else if (ind.sentido === "menos" && est >= 50) add("trampa", 3, `Estándar «≤ ${nfmt(est)} %» en un indicador de fallos: estás aceptando que falle la mitad o más. ¿No será el sentido contrario (≥ ${nfmt(est)} % de casos correctos)?`);
      else if (ind.sentido === "mas" && est < 50) add("trampa", 3, `Estándar «≥ ${nfmt(est)} %» es muy poco exigente para un indicador de cumplimiento. ¿Mide fallos? Entonces el sentido es «menos es mejor».`);
      if (ind.sentido === "mas" && est === 100) add("nota", 3, "Un estándar del 100 % no deja margen: un solo fallo ya lo incumple. Es razonable si el criterio es crítico; si no, valora un 95 %.");
    }
    if (umb !== null && !Number.isNaN(umb) && est !== null && !Number.isNaN(est) && ind.sentido) {
      if (ind.sentido === "mas" && umb >= est) add("error", 3, "El umbral de alarma tiene que estar por debajo del estándar (si más es mejor).");
      if (ind.sentido === "menos" && umb <= est) add("error", 3, "El umbral de alarma tiene que estar por encima del estándar (si menos es mejor).");
    }
    if (c === "centinela" && est !== null && est !== 0) add("trampa", 3, "En un centinela el estándar es siempre 0 casos.");
    if (usaDenominador(c) && lleno(ind.numerador) && !lleno(ind.denominador)) add("error", 3, "Falta el denominador: sin él no se puede calcular.");
    if (c === "porcentaje" && lleno(ind.numerador) && lleno(ind.denominador)) add("nota", 3, "Comprueba que los casos del numerador son una parte de los del denominador (los que cumplen, de entre todos los evaluados).");
    if (lleno(ind.numerador) && !lleno(ind.terminos)) add("trampa", 3, "Define los términos que se pueden entender de dos formas («correctamente», «completa», «en las primeras 24 h»: ¿desde qué hora?). Si dos auditoras los interpretan distinto, el indicador no es fiable.");
    return A;
  }

  // ---------- Paso 4 · Calidad ----------
  // Pistas automáticas a partir de lo ya escrito (no sustituyen al juicio de quien lo diseña)
  function pistasCalidad(ind) {
    const P = {}, falta = [];
    const def = ind.calculo === "centinela" || ind.calculo === "absoluto" ? lleno(ind.numerador) : lleno(ind.numerador) && lleno(ind.denominador);
    P.E = lleno(ind.nombre) && def ? { ok: true, t: "Tiene nombre y define qué casos se cuentan." } : { ok: false, t: "Falta el nombre o el numerador/denominador (paso 3)." };
    P.M = lleno(ind.calculo) && def && (lleno(ind.estandar) || ind.calculo === "centinela") ? { ok: true, t: "Tiene fórmula y estándar con el que compararse." } : { ok: false, t: "Falta la forma de cálculo, la fórmula o el estándar." };
    P.A = (ind.fuentes || []).length || lleno(ind.fuenteOtra) ? { ok: true, t: "La fuente de datos está identificada." } : { ok: false, t: "No has indicado de dónde salen los datos." };
    P.R = lleno(ind.hallazgo) || lleno(ind.criterio) ? { ok: true, t: "Está ligado a un hallazgo o criterio" + (ipr(ind.g, ind.o, ind.d) ? ` (IPR ${ipr(ind.g, ind.o, ind.d)}).` : ".") } : { ok: false, t: "No está ligado a ningún hallazgo del Mapa de Riesgos." };
    P.T = lleno(ind.periodicidad) ? { ok: true, t: "Periodicidad: " + ind.periodicidad.toLowerCase() + "." } : { ok: false, t: "Falta la periodicidad (paso 3)." };
    P.valido = lleno(ind.criterio) && lleno(ind.justTipo) ? { ok: true, t: "Sale de un criterio y has justificado el tipo." } : { ok: false, t: "Pregúntate si mide de verdad lo que dice medir." };
    P.fiable = lleno(ind.terminos) ? { ok: true, t: "Has definido los términos ambiguos." } : { ok: false, t: "Sin definir los términos ambiguos, dos auditoras pueden discrepar." };
    P.util = lleno(ind.responsable) ? { ok: true, t: "Tiene responsable: alguien puede actuar." } : { ok: false, t: "Falta el responsable: ¿quién puede mejorarlo?" };
    P.comparable = lleno(ind.unidad) || ["porcentaje", "compuesto"].includes(ind.calculo) ? { ok: true, t: "Unidad clara: se puede comparar en el tiempo y entre unidades." } : { ok: false, t: "Fija la unidad y el método para poder comparar." };
    Object.keys(P).forEach(k => { if (!P[k].ok) falta.push(k); });
    return P;
  }
  // Semáforo de la valoración que hace la persona (sí / dudoso / no)
  function estadoCalidad(ind) {
    const cal = ind.calidad || {};
    const v = CALIDAD.map(q => (cal[q.k] && cal[q.k].v) || "");
    const no = v.filter(x => x === "no").length, dud = v.filter(x => x === "dudoso").length, si = v.filter(x => x === "si").length;
    const contestadas = no + dud + si;
    let color = "gris", texto = "Sin valorar";
    if (no) { color = "rojo"; texto = "Rehacer"; }
    else if (contestadas === CALIDAD.length && !dud) { color = "verde"; texto = "Listo"; }
    else if (contestadas) { color = "ambar"; texto = "Revisar"; }
    return { color, texto, si, dudoso: dud, no, contestadas, total: CALIDAD.length };
  }
  // % de acuerdo del pilotaje (el kappa lo calcula Tamaño Muestral)
  function acuerdoPiloto(p) {
    const v = ["a", "b", "c", "d"].map(k => num(p && p[k]));
    if (v.some(x => x === null || Number.isNaN(x) || x < 0 || !Number.isInteger(x))) return null;
    const n = v[0] + v[1] + v[2] + v[3]; if (!n) return null;
    return { n, acuerdo: (v[0] + v[3]) / n, discrepan: v[1] + v[2] };
  }

  // ---------- Progreso del asistente ----------
  function pasosCompletos(ind) {
    const def = !usaDenominador(ind.calculo) ? lleno(ind.numerador) : lleno(ind.numerador) && lleno(ind.denominador);
    return {
      1: lleno(ind.nombre) && (lleno(ind.hallazgo) || lleno(ind.criterio)),
      2: lleno(ind.tipo) && lleno(ind.calculo),
      3: def && (lleno(ind.estandar) || ind.calculo === "centinela") && lleno(ind.periodicidad),
      4: estadoCalidad(ind).contestadas === CALIDAD.length,
      5: lleno(ind.objetivo)
    };
  }

  // ---------- Paso 5 · Muestra y enlaces a Tamaño Muestral ----------
  // Qué recomienda la app según la forma de cálculo
  function recomendacionMuestra(ind) {
    switch (ind.calculo) {
      case "centinela": return { modo: "sin", texto: "No se muestrea: se vigilan todos los casos." };
      case "absoluto": return { modo: "sin", texto: "Normalmente no se muestrea: se cuentan todos los eventos del periodo." };
      case "media": return { modo: "media", texto: "Para estimar una media hace falta una idea de la dispersión (desviación típica)." };
      case "compuesto": return { modo: "compuesto", texto: "Se revisan pacientes y se cuentan oportunidades de cada criterio." };
      case "porcentaje": case "tasa": case "ratio": return { modo: "elegir", texto: "¿Quieres estimar el valor o decidir si hay un problema?" };
      default: return { modo: "", texto: "Elige primero la forma de cálculo (paso 2)." };
    }
  }
  // Para LQAS, el estándar y el umbral se expresan siempre como % de casos correctos
  function lqasParams(ind) {
    const est = num(ind.estandar), umb = num(ind.umbral);
    if (est === null || umb === null || Number.isNaN(est) || Number.isNaN(umb)) return null;
    const inv = ind.sentido === "menos";
    return { pU: inv ? 100 - est : est, pL: inv ? 100 - umb : umb, invertido: inv };
  }
  function enlaceTM(calc, datos) {
    const limpio = {};
    Object.keys(datos).forEach(k => {
      const v = datos[k];
      if (k === "rows") limpio.rows = v;
      else if (v !== null && v !== undefined && String(v).trim() !== "") limpio[k] = String(v).trim().replace(",", ".");
    });
    return TM_URL + "#/" + calc + "?d=" + encodeURIComponent(JSON.stringify(limpio));
  }
  function enlacesTM(ind) {
    const m = ind.muestra || {}, L = {};
    L.proporcion = enlaceTM("proporcion", { conf: m.conf, e: m.e, p: m.p || "50", N: m.N, loss: m.loss });
    const lp = lqasParams(ind);
    if (lp) L.lqas = enlaceTM("lqas", { pU: nfmt(lp.pU).replace(",", "."), pL: nfmt(lp.pL).replace(",", "."), alpha: m.alpha, beta: m.beta, N: m.N, loss: m.loss });
    L.media = enlaceTM("media", { conf: m.conf, sigma: m.sigma, e: m.eMedia, N: m.N, loss: m.loss });
    const p = ind.piloto || {};
    if (["a", "b", "c", "d"].some(k => lleno(p[k]))) L.concordancia = enlaceTM("concordancia", { a: p.a, b: p.b, c: p.c, d: p.d });
    else L.concordancia = TM_URL + "#/concordancia";
    return L;
  }

  // ---------- Indicador compuesto (Saturno, UT 19) ----------
  function calcCompuesto(rows) {
    const errores = [], filas = [];
    (rows || []).forEach((r, i) => {
      const cum = num(r.cum), opo = num(r.opo);
      if (cum === null && opo === null && !lleno(r.nom)) return;
      const nom = lleno(r.nom) ? r.nom.trim() : `Criterio ${i + 1}`;
      if (cum === null || opo === null || Number.isNaN(cum) || Number.isNaN(opo) || !Number.isInteger(cum) || !Number.isInteger(opo) || cum < 0 || opo <= 0) {
        errores.push(`«${nom}»: escribe cumplimientos y oportunidades como números enteros (oportunidades mayor que 0).`); return;
      }
      if (cum > opo) { errores.push(`«${nom}»: los cumplimientos no pueden superar a las oportunidades.`); return; }
      filas.push({ i, nom, cum, opo, pct: cum / opo * 100 });
    });
    if (errores.length || filas.length < 2) return { valido: false, errores: errores.length ? errores : ["Hacen falta al menos dos criterios completos."], filas };
    const cum = filas.reduce((s, f) => s + f.cum, 0), opo = filas.reduce((s, f) => s + f.opo, 0);
    filas.forEach(f => { f.peso = f.opo / opo * 100; });
    const pct = cum / opo * 100, mediaPct = filas.reduce((s, f) => s + f.pct, 0) / filas.length;
    const flojo = filas.reduce((a, f) => f.pct < a.pct ? f : a, filas[0]);
    const pesado = filas.reduce((a, f) => f.opo > a.opo ? f : a, filas[0]);
    return { valido: true, errores: [], filas, cum, opo, pct, mediaPct, flojo, pesado };
  }
  const pct1 = x => (Math.round(x * 10) / 10).toFixed(1).replace(".", ",") + " %";
  function fraseCompuesto(nombre, r) {
    if (!r || !r.valido) return "";
    const tit = lleno(nombre) ? nombre.trim() : "El compuesto";
    let f = `${tit}: ${pct1(r.pct)} de las ${r.opo} oportunidades revisadas`;
    if (r.flojo.pct < r.pct) f += `; el punto débil es «${r.flojo.nom}» (${pct1(r.flojo.pct)}).`;
    else f += ".";
    return f;
  }

  // ---------- Fase 2 · Resultados por periodo (Verificar del PDCA) ----------
  // Cada medición: { periodo, num, den, valor, nota }. «valor» solo se usa si no hay numerador/denominador
  // (dato ya calculado) o en la media. Esta app no calcula LQAS: aplica la regla n/c anotada en el paso 5.
  const nuevaMedicion = (periodo = "") => ({ periodo, num: "", den: "", valor: "", nota: "" });
  // Recuentos con puntos de millar («2.950 estancias» = 2950)
  function numMiles(s) {
    const t = String(s == null ? "" : s).trim();
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) return Number(t.replace(/\./g, "").replace(",", "."));
    return num(t);
  }
  const esLqas = ind => ind.objetivo === "decidir";
  const conMuestra = ind => ["estimar", "decidir", "media"].includes(ind.objetivo);
  // Etiquetas de los campos según la forma de cálculo
  function camposMedicion(ind) {
    const c = ind.calculo;
    if (esLqas(ind)) return { num: "Numerador", den: "Revisados (n)", valor: false };
    if (c === "centinela" || c === "absoluto") return { num: "Nº de casos", den: false, valor: false };
    if (c === "media") return { num: false, den: "Casos medidos", valor: "Media", valorOblig: true };
    if (c === "tasa") return { num: "Casos", den: "Denominador", valor: "Ya calculada" };
    if (c === "porcentaje" || c === "compuesto") return { num: c === "compuesto" ? "Σ cumplimientos" : "Numerador", den: c === "compuesto" ? "Σ oportunidades" : "Denominador", valor: "% ya calculado" };
    return { num: "Numerador", den: "Denominador", valor: "Ya calculado" };
  }
  // Decimales con los que se muestra (y se compara con el estándar) cada forma de cálculo
  const decimales = ind => ["centinela", "absoluto"].includes(ind.calculo) ? 0 : ["tasa", "ratio"].includes(ind.calculo) ? 2 : 1;
  const redondea = (x, d) => Math.round(x * 10 ** d) / 10 ** d;
  const malNum = x => x === null || Number.isNaN(x);
  // Valora una medición: valor, semáforo (verde · rojo · gris) y motivo
  function evaluarMedicion(ind, m) {
    m = m || {};
    const out = { vacio: false, error: "", valor: null, estado: "gris", motivo: "", lqas: null };
    const num_ = numMiles(m.num), den_ = numMiles(m.den), val_ = num(m.valor);
    if ([m.num, m.den, m.valor].every(v => !lleno(v))) { out.vacio = true; out.motivo = "Sin datos"; return out; }
    if ([num_, den_, val_].some(x => Number.isNaN(x))) { out.error = "Escribe solo números."; return out; }
    const c = ind.calculo, est = num(ind.estandar);
    // LQAS: se aplica la regla anotada (verde si los casos correctos llegan a c)
    if (esLqas(ind)) {
      const n = num(ind.plan && ind.plan.n), cc = num(ind.plan && ind.plan.c);
      if (malNum(n) || malNum(cc) || !Number.isInteger(n) || !Number.isInteger(cc) || n <= 0 || cc > n) { out.error = "Anota primero n y c del plan LQAS en el paso 5."; return out; }
      const rev = den_ === null ? n : den_;
      if (num_ === null) { out.error = "Escribe el numerador (cuántos casos lo cumplen" + (ind.sentido === "menos" ? ", o cuántos fallan si el indicador cuenta fallos" : "") + ")."; return out; }
      if (!Number.isInteger(num_) || !Number.isInteger(rev) || num_ < 0 || rev <= 0) { out.error = "Escribe números enteros (0, 1, 2…)."; return out; }
      if (num_ > rev) { out.error = `El numerador (${num_}) no puede superar a los revisados (${rev}).`; return out; }
      const correctos = ind.sentido === "menos" ? rev - num_ : num_, fallos = rev - correctos;
      out.valor = num_ / rev * 100;
      out.lqas = { n, c: cc, rev, correctos, fallos, faltan: Math.max(0, n - rev) };
      if (rev > n) { out.motivo = `El plan es revisar ${n}: aplica la regla solo a los ${n} primeros del sorteo.`; return out; }
      if (correctos >= cc) { out.estado = "verde"; out.motivo = `${correctos} de ${rev} correctos; la regla pide al menos ${cc} de ${n}.` + (rev < n ? " Ya es verde aunque falten casos." : ""); }
      else if (fallos > n - cc) { out.estado = "rojo"; out.motivo = `${correctos} de ${rev} correctos; la regla pide al menos ${cc} de ${n}.` + (rev < n ? " Ya es rojo aunque falten casos." : ""); }
      else out.motivo = `Faltan ${n - rev} casos por revisar: la regla solo vale con los ${n} del plan.`;
      return out;
    }
    let v = null;
    if (c === "centinela" || c === "absoluto") {
      if (num_ === null) { out.error = "Escribe el número de casos."; return out; }
      if (!Number.isInteger(num_) || num_ < 0) { out.error = "El número de casos es un entero (0, 1, 2…)."; return out; }
      v = num_;
    } else if (c === "media") {
      if (val_ === null) { out.error = "Escribe la media del periodo."; return out; }
      v = val_;
    } else if (num_ !== null || den_ !== null) {
      if (num_ === null || den_ === null) { out.error = "Faltan el numerador o el denominador."; return out; }
      if (num_ < 0 || den_ <= 0) { out.error = "El denominador tiene que ser mayor que 0."; return out; }
      if ((c === "porcentaje" || c === "compuesto") && num_ > den_) { out.error = `El numerador (${num_}) no puede superar al denominador (${den_}).`; return out; }
      const base = c === "tasa" ? (lleno(ind.base) ? numMiles(ind.base) : 1000) : 1;
      if (malNum(base) || base <= 0) { out.error = "Revisa la base de la tasa en el paso 3 (p. ej. 1.000)."; return out; }
      v = c === "porcentaje" || c === "compuesto" ? num_ / den_ * 100 : c === "tasa" ? num_ / den_ * base : num_ / den_;
    } else v = val_;
    if ((c === "porcentaje" || c === "compuesto") && (v < 0 || v > 100)) { out.error = "Un porcentaje va de 0 a 100."; return out; }
    out.valor = v;
    if (c === "centinela") { out.estado = v === 0 ? "verde" : "rojo"; out.motivo = v === 0 ? "Ningún caso." : `${v} caso${v === 1 ? "" : "s"}: analiza cada uno (análisis causa raíz).`; return out; }
    if (!ind.sentido || malNum(est)) { out.motivo = "Sin estándar o sin sentido (paso 3): no se puede comparar."; return out; }
    const vr = redondea(v, decimales(ind)), cumple = ind.sentido === "menos" ? vr <= est : vr >= est;
    out.estado = cumple ? "verde" : "rojo";
    out.motivo = cumple ? "Cumple el estándar." : "No cumple el estándar.";
    return out;
  }
  // Texto del valor con su unidad: «92,5 %», «0,65 por 1.000 estancias», «3 casos»
  function textoValor(ind, v) {
    if (!Number.isFinite(v)) return "—";
    const u = unidad(ind), dec = decimales(ind);
    const t = redondea(v, dec).toFixed(dec).replace(".", ",");
    return t + (u === "%" ? " %" : u ? " " + u : "");
  }
  // Lo que se enseña como «valor» de una medición: en LQAS, los casos correctos (no un %)
  function textoResultado(ind, f) {
    if (f.lqas) return `${f.lqas.correctos} de ${f.lqas.rev} correctos`;
    return f.valor === null ? "—" : textoValor(ind, f.valor);
  }
  // Serie completa y resumen de la tendencia (sin pruebas estadísticas)
  function analizarResultados(ind) {
    const filas = (ind.resultados || []).map((m, i) => Object.assign({ i, periodo: lleno(m.periodo) ? m.periodo.trim() : `Periodo ${i + 1}` }, evaluarMedicion(ind, m)));
    const validas = filas.filter(f => !f.vacio && !f.error && f.valor !== null);
    const juzgadas = validas.filter(f => f.estado !== "gris");
    const frases = [];
    const ult = juzgadas.slice(-4);
    if (ult.length) {
      const v = ult.filter(f => f.estado === "verde").length;
      frases.push(`${esLqas(ind) ? "Verde" : "Cumple el estándar"} en ${v} de ${ult.length === juzgadas.length ? "los " + ult.length : "los últimos " + ult.length} periodo${ult.length === 1 ? "" : "s"}.`);
      let racha = 0; for (let k = juzgadas.length - 1; k >= 0 && juzgadas[k].estado === "rojo"; k--) racha++;
      if (racha >= 2) frases.push(`${racha} periodos seguidos en rojo: toca actuar (analizar causas y planificar mejoras).`);
      else if (racha === 1) frases.push("El último periodo está en rojo: busca qué ha fallado antes de que se repita.");
    }
    let cambio = "";
    if (!esLqas(ind) && validas.length >= 2 && ind.sentido) {
      const a = validas[validas.length - 2].valor, b = validas[validas.length - 1].valor;
      const mejor = ind.sentido === "menos" ? b < a : b > a;
      cambio = redondea(a, decimales(ind)) === redondea(b, decimales(ind)) ? "igual" : mejor ? "mejora" : "empeora";
      frases.push(cambio === "igual" ? "Igual que el periodo anterior." : `${cambio === "mejora" ? "Mejora" : "Empeora"} respecto al periodo anterior (${textoValor(ind, a)} → ${textoValor(ind, b)}).`);
    }
    if (validas.length && validas.length < 3) frases.push("Con menos de tres periodos todavía no se puede hablar de tendencia.");
    return { filas, validas, ultima: validas[validas.length - 1] || null, cambio, frases };
  }
  // Eje vertical del gráfico: mínimo, máximo y marcas «redondas»
  function ejeY(valores, estandar, esPct) {
    const vs = valores.filter(Number.isFinite).concat(Number.isFinite(estandar) ? [estandar] : []);
    if (!vs.length) return { min: 0, max: esPct ? 100 : 1, marcas: esPct ? [0, 25, 50, 75, 100] : [0, 1] };
    let lo = Math.min(...vs), hi = Math.max(...vs);
    if (hi === lo) { hi = hi + (hi === 0 ? 1 : Math.abs(hi) * 0.1); lo = lo - (lo === 0 ? 0 : Math.abs(lo) * 0.1); }
    const pad = (hi - lo) * 0.15; lo -= pad; hi += pad;
    if (Math.min(...vs) >= 0) lo = Math.max(0, lo);
    if (esPct) { lo = Math.max(0, lo); hi = Math.min(100, hi); }
    const bruto = (hi - lo) / 4, pot = 10 ** Math.floor(Math.log10(bruto));
    const paso = [1, 2, 2.5, 5, 10].map(k => k * pot).find(k => k >= bruto) || 10 * pot;
    const min = Math.floor(lo / paso) * paso, max = Math.ceil(hi / paso) * paso, marcas = [];
    for (let x = min; x <= max + paso / 1e6; x += paso) marcas.push(Math.round(x * 1e6) / 1e6);
    return { min, max, marcas };
  }
  // Periodo siguiente para el botón «Añadir periodo»: 2026-T4 → 2027-T1, Ene 2026 → Feb 2026, 2026 → 2027
  const MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
  function siguientePeriodo(prev) {
    const p = String(prev || "").trim();
    let m;
    if ((m = /^(\d{4})-([TSQ])(\d)$/i.exec(p))) {
      const max = { T: 4, S: 2, Q: 2 }[m[2].toUpperCase()], k = Number(m[3]);
      return k >= max ? `${Number(m[1]) + 1}-${m[2]}1` : `${m[1]}-${m[2]}${k + 1}`;
    }
    if ((m = /^(\d{4})$/.exec(p))) return String(Number(m[1]) + 1);
    if ((m = /^([a-záéíóú]+)\.?\s+(\d{4})$/i.exec(p))) {
      const i = MESES.indexOf(sinTildes(m[1]).slice(0, 3)); if (i < 0) return "";
      const may = m[1][0] === m[1][0].toUpperCase(), sig = MESES[(i + 1) % 12];
      return `${may ? sig[0].toUpperCase() + sig.slice(1) : sig} ${i === 11 ? Number(m[2]) + 1 : m[2]}`;
    }
    if ((m = /^(.*?)(\d+)$/.exec(p))) return m[1] + (Number(m[2]) + 1);
    return "";
  }
  function ejemploPeriodo(periodicidad) {
    return { "Mensual": "Ene 2026", "Trimestral": "2026-T1", "Semestral": "2026-S1", "Anual": "2026", "Semanal": "Semana 1" }[periodicidad] || "p. ej. 2026-T1";
  }
  // Enlace a «Ya tengo los resultados» de Tamaño Muestral con los datos de la medición
  function enlaceAnalisis(ind, m) {
    const ev = evaluarMedicion(ind, m), mm = ind.muestra || {};
    if (esLqas(ind)) {
      const lp = lqasParams(ind); if (!lp || !ev.lqas) return "";
      return enlaceTM("lqas", { pU: nfmt(lp.pU).replace(",", "."), pL: nfmt(lp.pL).replace(",", "."), alpha: mm.alpha, beta: mm.beta, N: mm.N, loss: mm.loss, an_rev: ev.lqas.rev, an_cum: ev.lqas.correctos });
    }
    if (ind.calculo !== "porcentaje" || ev.error || ev.vacio) return "";
    const n_ = numMiles(m.num), d_ = numMiles(m.den);
    if (n_ === null || d_ === null || !Number.isInteger(n_) || !Number.isInteger(d_)) return "";
    const est = num(ind.estandar), inv = ind.sentido === "menos";
    return enlaceTM("proporcion", { conf: mm.conf, e: mm.e, p: mm.p || "50", N: mm.N, loss: mm.loss, an_rev: d_, an_cum: inv ? d_ - n_ : n_, an_est: malNum(est) || est <= 0 || est >= 100 ? "" : nfmt(inv ? 100 - est : est).replace(",", ".") });
  }

  // ---------- Datos: validación de lo que se importa o llega por enlace ----------
  const CAMPOS_TEXTO = ["subproceso", "hallazgo", "g", "o", "d", "criterio", "nombre", "tipo", "justTipo", "calculo",
    "numerador", "denominador", "base", "unidad", "poblacion", "exclusiones", "fuenteOtra", "sentido", "estandar", "umbral",
    "periodicidad", "responsable", "terminos", "observaciones", "objetivo", "creado", "modificado"];
  const ENUMS = { tipo: Object.keys(TIPOS), calculo: Object.keys(CALCULOS), sentido: ["mas", "menos"], objetivo: ["estimar", "decidir", "media", "compuesto", "sin"] };
  const txt = (v, max) => (typeof v === "string" || typeof v === "number") ? String(v).slice(0, max) : "";
  const idValido = v => typeof v === "string" && /^[a-z0-9-]{1,40}$/i.test(v);

  function sanearIndicador(o, idNuevo, fecha) {
    if (!o || typeof o !== "object") return null;
    const ind = nuevoIndicador(idValido(o.id) ? o.id : idNuevo, fecha);
    CAMPOS_TEXTO.forEach(k => { if (o[k] !== undefined) ind[k] = txt(o[k], 2000); });
    Object.keys(ENUMS).forEach(k => { if (!ENUMS[k].includes(ind[k])) ind[k] = ""; });
    if (Array.isArray(o.fuentes)) ind.fuentes = o.fuentes.slice(0, 20).map(f => txt(f, 100)).filter(Boolean);
    if (o.calidad && typeof o.calidad === "object") CALIDAD.forEach(q => {
      const c = o.calidad[q.k]; if (!c || typeof c !== "object") return;
      ind.calidad[q.k] = { v: ["si", "dudoso", "no"].includes(c.v) ? c.v : "", t: txt(c.t, 1000) };
    });
    ["piloto", "muestra", "plan"].forEach(g => {
      if (o[g] && typeof o[g] === "object") Object.keys(ind[g]).forEach(k => { if (o[g][k] !== undefined) ind[g][k] = txt(o[g][k], g === "plan" ? 500 : 20); });
    });
    if (Array.isArray(o.resultados)) ind.resultados = o.resultados.slice(0, 200).filter(m => m && typeof m === "object")
      .map(m => ({ periodo: txt(m.periodo, 60), num: txt(m.num, 20), den: txt(m.den, 20), valor: txt(m.valor, 20), nota: txt(m.nota, 500) }));
    return ind;
  }
  function sanearCompuesto(o, idNuevo, fecha) {
    if (!o || typeof o !== "object") return null;
    const c = nuevoCompuesto(idValido(o.id) ? o.id : idNuevo, fecha);
    ["nombre", "pacientes", "estandar", "conclusion", "creado", "modificado"].forEach(k => { if (o[k] !== undefined) c[k] = txt(o[k], k === "conclusion" ? 2000 : 300); });
    if (Array.isArray(o.rows)) c.rows = o.rows.slice(0, 50).map(r => ({ nom: txt(r && r.nom, 300), cum: txt(r && r.cum, 20), opo: txt(r && r.opo, 20), ind: idValido(r && r.ind) ? r.ind : "" }));
    if (!c.rows.length) c.rows = nuevoCompuesto("x", fecha).rows;
    return c;
  }
  // Copia de seguridad completa: { app, v, exportado, indicadores: [], compuestos: [] }
  function sanearDatos(o, genId, fecha) {
    if (!o || typeof o !== "object" || !Array.isArray(o.indicadores)) return null;
    const indicadores = o.indicadores.slice(0, 500).map(x => sanearIndicador(x, genId(), fecha)).filter(Boolean);
    const compuestos = Array.isArray(o.compuestos) ? o.compuestos.slice(0, 200).map(x => sanearCompuesto(x, genId(), fecha)).filter(Boolean) : [];
    return { indicadores, compuestos };
  }
  // Fusiona sin duplicar: si un id ya existe se conserva el más reciente
  function fusionar(actual, nuevo) {
    const une = (a, b) => {
      const m = new Map(a.map(x => [x.id, x]));
      b.forEach(x => { const y = m.get(x.id); if (!y || String(x.modificado) > String(y.modificado)) m.set(x.id, x); });
      return Array.from(m.values());
    };
    return { indicadores: une(actual.indicadores, nuevo.indicadores), compuestos: une(actual.compuestos, nuevo.compuestos) };
  }

  // ---------- Exportar a Excel (definición de hojas para Xlsx.construirLibroXlsx) ----------
  const ESTILOS_XLSX = {
    titulo: { negrita: true, tamano: 14, colorTexto: "FFFFFF", colorFondo: "0B6F66", borde: false },
    sub: { cursiva: true, colorTexto: "444444", borde: false },
    cab: { negrita: true, colorTexto: "FFFFFF", colorFondo: "0D1B2A", alineacion: "center", ajustarTexto: true },
    celda: { ajustarTexto: true, vertical: "top" },
    centro: { alineacion: "center", vertical: "top", ajustarTexto: true },
    total: { negrita: true, colorFondo: "E6F3F7" },
    flojo: { colorFondo: "FBF0DC", ajustarTexto: true },
    verde: { colorFondo: "DCFCE7", colorTexto: "14532D", negrita: true, alineacion: "center", vertical: "top", ajustarTexto: true },
    rojo: { colorFondo: "FFE4E6", colorTexto: "881337", negrita: true, alineacion: "center", vertical: "top", ajustarTexto: true }
  };
  const fuentesTexto = ind => [...(ind.fuentes || []), ind.fuenteOtra].filter(lleno).join(", ");
  const calTexto = (ind, k) => { const c = (ind.calidad || {})[k] || {}; const v = { si: "Sí", dudoso: "Dudoso", no: "No" }[c.v] || ""; return [v, c.t].filter(lleno).join(". "); };
  function hojasExcel(datos, fecha) {
    const T = (texto, estilo = "celda") => ({ texto: texto == null ? "" : String(texto), estilo });
    const inds = datos.indicadores;
    const cab5A = ["Nº", "Subproceso", "Hallazgo", "IPR", "Criterio", "Indicador", "Tipo", "Forma de cálculo", "Justificación", "Específico", "Medible", "Alcanzable", "Relevante", "Temporal", "Válido", "Fiable", "Útil", "Comparable"];
    const h5A = [[T("Caracterización de indicadores", "titulo")], [T("Exportado de la app Indicadores · " + fecha, "sub")], cab5A.map(x => T(x, "cab"))];
    inds.forEach((ind, i) => h5A.push([T(i + 1, "centro"), T(ind.subproceso), T(ind.hallazgo), T(ipr(ind.g, ind.o, ind.d) || "", "centro"), T(ind.criterio), T(ind.nombre),
      T(TIPOS[ind.tipo] ? TIPOS[ind.tipo].nom : ""), T(CALCULOS[ind.calculo] ? CALCULOS[ind.calculo].nom : ""), T(ind.justTipo),
      ...CALIDAD.map(q => T(calTexto(ind, q.k)))]));
    h5A[0][0].combinar = cab5A.length;
    const cab5B = ["Nº", "Subproceso", "Hallazgo", "Indicador", "Numerador", "Denominador", "Fórmula", "Unidad", "Población y exclusiones", "Fuente de datos", "Sentido", "Estándar", "Umbral de alarma", "Periodicidad", "Responsable", "Definición de términos", "Observaciones", "Plan de muestreo"];
    const h5B = [[T("Definición operativa de indicadores", "titulo")], [T("Exportado de la app Indicadores · " + fecha, "sub")], cab5B.map(x => T(x, "cab"))];
    inds.forEach((ind, i) => h5B.push([T(i + 1, "centro"), T(ind.subproceso), T(ind.hallazgo), T(ind.nombre), T(ind.numerador), T(usaDenominador(ind.calculo) ? ind.denominador : "—"),
      T(formula(ind)), T(unidad(ind)), T([ind.poblacion, lleno(ind.exclusiones) ? "Excluye: " + ind.exclusiones : ""].filter(lleno).join(". ")), T(fuentesTexto(ind)),
      T(ind.sentido === "menos" ? "Menos es mejor" : ind.sentido === "mas" ? "Más es mejor" : ""), T(textoEstandar(ind)), T(textoEstandar(ind, "umbral")),
      T(ind.periodicidad), T(ind.responsable), T(ind.terminos), T(ind.observaciones), T(textoPlan(ind))]));
    h5B[0][0].combinar = cab5B.length;
    const hojas = [
      { nombre: "Caracterización", filas: h5A, opciones: { anchoColumnas: [5, 18, 22, 6, 28, 32, 12, 14, 28, 18, 18, 18, 18, 18, 18, 18, 18, 18], congelar: { filas: 3, columnas: 0 }, altoFilas: { 1: 24 } } },
      { nombre: "Definición operativa", filas: h5B, opciones: { anchoColumnas: [5, 18, 22, 32, 26, 26, 30, 12, 26, 22, 14, 14, 14, 13, 18, 30, 26, 24], congelar: { filas: 3, columnas: 0 }, altoFilas: { 1: 24 } } }
    ];
    if (datos.compuestos.length) {
      const hc = [[T("Indicadores compuestos", "titulo", 5)], [T("Compuesto = Σ cumplimientos ÷ Σ oportunidades × 100. Nunca sin su desglose.", "sub")]];
      hc[0][0].combinar = 5;
      datos.compuestos.forEach(c => {
        const r = calcCompuesto(c.rows);
        hc.push([]); hc.push([T(c.nombre || "Compuesto sin nombre", "total"), T("", "total"), T("", "total"), T("", "total"), T("", "total")]);
        hc.push(["Criterio", "Cumplimientos", "Oportunidades", "% cumplimiento", "Peso en el compuesto"].map(x => T(x, "cab")));
        if (!r.valido) { hc.push([T(r.errores.join(" "))]); return; }
        r.filas.forEach(f => hc.push([T(f.nom, f === r.flojo ? "flojo" : "celda"), { numero: f.cum, estilo: "centro" }, { numero: f.opo, estilo: "centro" }, T(pct1(f.pct), "centro"), T(pct1(f.peso), "centro")]));
        hc.push([T("Indicador compuesto", "total"), { numero: r.cum, estilo: "total" }, { numero: r.opo, estilo: "total" }, T(pct1(r.pct), "total"), T("100 %", "total")]);
        hc.push([T("Media de los % (solo para comparar)"), T(""), T(""), T(pct1(r.mediaPct), "centro"), T("")]);
        hc.push([T("Conclusión: " + (lleno(c.conclusion) ? c.conclusion : fraseCompuesto(c.nombre, r)))]);
        hc[hc.length - 1][0].combinar = 5;
      });
      hojas.push({ nombre: "Compuestos", filas: hc, opciones: { anchoColumnas: [46, 15, 15, 15, 20] } });
    }
    const conRes = inds.map((ind, i) => ({ ind, i, a: analizarResultados(ind) })).filter(x => x.a.filas.some(f => !f.vacio));
    if (conRes.length) {
      const cabR = ["Nº", "Indicador", "Estándar", "Periodo", "Numerador", "Denominador", "Valor", "Resultado", "Detalle", "Nota"];
      const hr = [[T("Resultados por periodo (Verificar del PDCA)", "titulo")], [T("Exportado de la app Indicadores · " + fecha + ". En LQAS el valor es solo el dato de la muestra: LQAS decide, no mide.", "sub")], cabR.map(x => T(x, "cab"))];
      hr[0][0].combinar = cabR.length;
      conRes.forEach(({ ind, i, a }) => a.filas.filter(f => !f.vacio).forEach(f => {
        const m = ind.resultados[f.i];
        const res = f.error ? "Revisar" : f.estado === "verde" ? (esLqas(ind) ? "Verde (LQAS)" : "Cumple") : f.estado === "rojo" ? (esLqas(ind) ? "Rojo (LQAS)" : "No cumple") : "Sin valorar";
        hr.push([T(i + 1, "centro"), T(ind.nombre), T(esLqas(ind) ? textoEstandar(ind) + " · n = " + (ind.plan.n || "?") + ", c = " + (ind.plan.c || "?") : textoEstandar(ind)), T(f.periodo),
          T(m.num, "centro"), T(m.den || (f.lqas ? f.lqas.rev : ""), "centro"), T(f.valor === null ? "" : textoResultado(ind, f), "centro"),
          T(res, f.error ? "centro" : f.estado === "verde" ? "verde" : f.estado === "rojo" ? "rojo" : "centro"), T(f.error || f.motivo), T(m.nota)]);
      }));
      hojas.push({ nombre: "Resultados", filas: hr, opciones: { anchoColumnas: [5, 40, 18, 14, 12, 13, 18, 14, 40, 30], congelar: { filas: 3, columnas: 0 }, altoFilas: { 1: 24 } } });
    }
    return { estilos: ESTILOS_XLSX, hojas };
  }
  function textoPlan(ind) {
    const p = ind.plan || {}, partes = [];
    if (ind.objetivo === "sin") return "Sin muestreo: todos los casos";
    if (lleno(p.n)) partes.push("n = " + p.n);
    if (lleno(p.c)) partes.push("c = " + p.c);
    if (lleno(p.nota)) partes.push(p.nota);
    const m = { estimar: "Estimar", decidir: "LQAS", media: "Estimar media", compuesto: "Compuesto" }[ind.objetivo];
    return (m ? m + (partes.length ? ": " : "") : "") + partes.join(" · ");
  }

  const API = {
    TM_URL, AUD_URL, AMFE_URL, FLUJO_URL, DAFO_URL, MAPA_URL, VERSION_DATOS, TIPOS, CALCULOS, CALIDAD, PERIODICIDADES, FUENTES,
    num, lleno, nfmt, sinTildes, nuevoIndicador, nuevoCompuesto, ipr, sugerirNombre, sentidoSugerido,
    unidadPorDefecto, unidad, formula, usaDenominador, textoEstandar, avisos, pistasCalidad, estadoCalidad, acuerdoPiloto,
    pasosCompletos, recomendacionMuestra, lqasParams, enlaceTM, enlacesTM, calcCompuesto, fraseCompuesto, pct1,
    nuevaMedicion, numMiles, esLqas, conMuestra, camposMedicion, evaluarMedicion, textoValor, textoResultado, analizarResultados, ejeY, siguientePeriodo, ejemploPeriodo, enlaceAnalisis,
    sanearIndicador, sanearCompuesto, sanearDatos, fusionar, hojasExcel, textoPlan, fuentesTexto
  };
  if (typeof module !== "undefined" && module.exports) module.exports = API; else raiz.Logica = API;
})(typeof window !== "undefined" ? window : globalThis);
