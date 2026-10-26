// Tests de logic.js y xlsx.js · node tests/logic.test.mjs
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const L = require("../logic.js");
const X = require("../xlsx.js");

let ok = 0, mal = 0;
function t(nombre, cond) { if (cond) ok++; else { mal++; console.error("✗ " + nombre); } }
const cerca = (a, b, tol = 1e-9) => Math.abs(a - b) < tol;

// IPR
t("IPR 10·7·7 = 490", L.ipr("10", "7", "7") === 490);
t("IPR incompleto = null", L.ipr("10", "", "7") === null);
t("IPR fuera de rango = null", L.ipr("11", "7", "7") === null);
t("IPR decimal = null", L.ipr("7,5", "7", "7") === null);

// Nombre sugerido y sentido
t("sugerirNombre", L.sugerirNombre("Se registra la valoración del riesgo de UPP en las primeras 24 h.") === "% de casos en los que se registra la valoración del riesgo de UPP en las primeras 24 h");
t("sentido: discrepancias → menos", L.sentidoSugerido("% de discrepancias entre el nivel asignado y el revisado") === "menos");
t("sentido: espera >30 min → menos", L.sentidoSugerido("% de pacientes que esperan más de 30 minutos") === "menos");
t("sentido: caídas → menos", L.sentidoSugerido("Caídas por cada 1.000 estancias") === "menos");
t("sentido: valoración registrada → más", L.sentidoSugerido("% de pacientes con valoración del riesgo de UPP registrada") === "mas");
t("sentido: vacío → null", L.sentidoSugerido("  ") === null);

// Fórmula y estándar
const ind = L.nuevoIndicador("a1", "2026-10-19");
Object.assign(ind, { calculo: "porcentaje", numerador: "pacientes con valoración registrada en 24 h", denominador: "pacientes ingresados revisados", sentido: "mas", estandar: "95" });
t("fórmula porcentaje", L.formula(ind) === "(Nº de pacientes con valoración registrada en 24 h ÷ Nº de pacientes ingresados revisados) × 100");
t("estándar ≥ 95 %", L.textoEstandar(ind) === "≥ 95 %");
const tasa = Object.assign(L.nuevoIndicador("a2", ""), { calculo: "tasa", numerador: "caídas", denominador: "estancias", base: "1.000" });
t("fórmula tasa", L.formula(tasa) === "(Nº de caídas ÷ estancias) × 1.000");
t("unidad tasa", L.unidad(tasa) === "por 1.000 estancias");
const cent = Object.assign(L.nuevoIndicador("a3", ""), { calculo: "centinela", numerador: "cirugía en lugar equivocado" });
t("estándar centinela", L.textoEstandar(cent).startsWith("0 casos"));

// Avisos: la trampa del sentido invertido (Anexo I de la auditoría modelo)
const inv = Object.assign(L.nuevoIndicador("a4", ""), { nombre: "% de discrepancias entre el nivel de prioridad asignado y el reevaluado", calculo: "porcentaje", sentido: "mas", estandar: "95", numerador: "discrepancias", denominador: "casos revisados" });
const av = L.avisos(inv);
t("detecta nombre negativo con «más es mejor»", av.some(a => a.nivel === "trampa" && /algo malo/.test(a.texto)));
inv.sentido = "menos";
t("detecta ≤ 95 % en indicador de fallos", L.avisos(inv).some(a => /falle la mitad/.test(a.texto)));
inv.estandar = "5";
t("sin trampa con ≤ 5 %", !L.avisos(inv).some(a => a.nivel === "trampa" && /mitad|algo malo/.test(a.texto)));
const um = Object.assign(L.nuevoIndicador("a5", ""), { calculo: "porcentaje", sentido: "mas", estandar: "85", umbral: "90" });
t("umbral por encima del estándar = error", L.avisos(um).some(a => a.nivel === "error" && /umbral/.test(a.texto)));
t("media maquilla", L.avisos(Object.assign(L.nuevoIndicador("x", ""), { calculo: "media" })).some(a => /maquillar/.test(a.texto)));
t("centinela no se muestrea", L.avisos(cent).some(a => /no se muestrea/.test(a.texto)));
t("falta denominador", L.avisos(Object.assign(L.nuevoIndicador("x", ""), { calculo: "porcentaje", numerador: "x" })).some(a => a.nivel === "error"));

// Calidad
const c1 = L.nuevoIndicador("c1", "");
t("calidad sin valorar", L.estadoCalidad(c1).color === "gris");
L.CALIDAD.forEach(q => { c1.calidad[q.k] = { v: "si", t: "" }; });
t("calidad todo sí = verde", L.estadoCalidad(c1).color === "verde");
c1.calidad.fiable.v = "dudoso";
t("un dudoso = ámbar", L.estadoCalidad(c1).color === "ambar");
c1.calidad.util.v = "no";
t("un no = rojo", L.estadoCalidad(c1).color === "rojo");
const ac = L.acuerdoPiloto({ a: "20", b: "2", c: "3", d: "5" });
t("acuerdo piloto 25/30", ac && ac.n === 30 && cerca(ac.acuerdo, 25 / 30));
t("acuerdo piloto incompleto", L.acuerdoPiloto({ a: "20", b: "", c: "3", d: "5" }) === null);

// Enlaces a Tamaño Muestral
const lq = Object.assign(L.nuevoIndicador("l1", ""), { calculo: "porcentaje", sentido: "mas", estandar: "85", umbral: "60" });
const urlL = L.enlacesTM(lq).lqas;
const dL = JSON.parse(decodeURIComponent(urlL.split("?d=")[1]));
// [eco]
t("LQAS URL base", urlL.startsWith("https://cjgaland.github.io/tamano-muestral/#/lqas?d="));
// [/eco]
/* [ind]
t("LQAS URL base (Tamaño Muestral incluido)", urlL.startsWith("tamano-muestral/#/lqas?d="));
[/ind] */
t("LQAS pU/pL como texto en %", dL.pU === "85" && dL.pL === "60" && dL.alpha === "5" && dL.beta === "20");
t("LQAS sin N vacío", !("N" in dL));
const lqi = Object.assign(L.nuevoIndicador("l2", ""), { calculo: "porcentaje", sentido: "menos", estandar: "5", umbral: "20" });
const dLi = JSON.parse(decodeURIComponent(L.enlacesTM(lqi).lqas.split("?d=")[1]));
t("LQAS indicador de fallos se invierte (≤5 % → 95, umbral 20 → 80)", dLi.pU === "95" && dLi.pL === "80");
t("LQAS sin umbral = sin enlace", !L.enlacesTM(Object.assign(L.nuevoIndicador("x", ""), { estandar: "90" })).lqas);
const dP = JSON.parse(decodeURIComponent(L.enlacesTM(lq).proporcion.split("?d=")[1]));
t("proporción con p = 50 por defecto", dP.p === "50" && dP.conf === "95" && dP.e === "10");
lq.piloto = { a: "20", b: "2", c: "3", d: "5" };
const dK = JSON.parse(decodeURIComponent(L.enlacesTM(lq).concordancia.split("?d=")[1]));
t("kappa con a, b, c, d", dK.a === "20" && dK.d === "5");

// Compuesto: ejemplo del infarto del tema 3 (28/40 = 70 %; media de % = 63,9 %)
const inf = [["Aspirina 24 h", 10, 12], ["β-bloq. 24 h", 4, 6], ["Aspirina alta", 4, 5], ["β-bloq. alta", 2, 4], ["Antitabaco", 1, 3], ["Trombolíticos", 7, 10]].map(([nom, cum, opo]) => ({ nom, cum: String(cum), opo: String(opo) }));
const r = L.calcCompuesto(inf);
t("compuesto válido", r.valido);
t("compuesto 28/40 = 70 %", r.cum === 28 && r.opo === 40 && cerca(r.pct, 70));
t("media de % = 63,9 %", Math.round(r.mediaPct * 10) / 10 === 63.9);
t("punto débil = antitabaco", r.flojo.nom === "Antitabaco");
t("más peso = aspirina 24 h (30 %)", r.pesado.nom === "Aspirina 24 h" && cerca(r.filas[0].peso, 30));
t("frase titular + desglose", /70,0 % de las 40 oportunidades.*Antitabaco/.test(L.fraseCompuesto("Atención al infarto", r)));
t("compuesto cum > opo = error", !L.calcCompuesto([{ nom: "a", cum: "5", opo: "4" }, { nom: "b", cum: "1", opo: "2" }]).valido);
t("compuesto con un solo criterio no vale", !L.calcCompuesto([{ nom: "a", cum: "3", opo: "4" }]).valido);
t("compuesto ignora filas vacías", L.calcCompuesto([...inf, { nom: "", cum: "", opo: "" }]).valido);

// Validación de datos importados
let n = 0; const gen = () => "g" + (++n);
const sucio = { indicadores: [{ id: "ok-1", nombre: "<img src=x onerror=alert(1)>", tipo: "hackeo", calculo: "porcentaje", calidad: { E: { v: "si", t: "x" }, Z: { v: "si" } }, raro: "x", fuentes: ["Historia clínica", 7, { a: 1 }] }, "basura", { id: "../../x" }], compuestos: [{ nombre: "C", rows: [{ nom: "a", cum: 1, opo: 2 }] }] };
const s = L.sanearDatos(sucio, gen, "2026-10-19");
t("sanea: conserva válidos", s.indicadores.length === 2 && s.compuestos.length === 1);
t("sanea: texto se guarda tal cual (se escapa al pintar)", s.indicadores[0].nombre.startsWith("<img"));
t("sanea: enum inválido → vacío", s.indicadores[0].tipo === "" && s.indicadores[0].calculo === "porcentaje");
t("sanea: claves desconocidas fuera", !("raro" in s.indicadores[0]) && !("Z" in s.indicadores[0].calidad));
t("sanea: id inválido → nuevo", s.indicadores[1].id === "g2" || s.indicadores[1].id.startsWith("g"));
t("sanea: fuentes solo texto", s.indicadores[0].fuentes.length === 2 && s.indicadores[0].fuentes[1] === "7");
t("sanea: no es copia → null", L.sanearDatos({ foo: 1 }, gen, "") === null);
const f = L.fusionar({ indicadores: [{ id: "a", modificado: "2026-10-01" }], compuestos: [] }, { indicadores: [{ id: "a", modificado: "2026-10-05", nombre: "nuevo" }, { id: "b", modificado: "" }], compuestos: [] });
t("fusionar: conserva el más reciente y añade nuevos", f.indicadores.length === 2 && f.indicadores.find(x => x.id === "a").nombre === "nuevo");

// Excel
const datos = { indicadores: [ind, tasa, cent], compuestos: [{ nombre: "Infarto", rows: inf, conclusion: "" }] };
const libro = L.hojasExcel(datos, "19/10/2026");
t("Excel: 3 hojas", libro.hojas.length === 3);
t("Excel: 5A tiene 3 filas de cabecera + 3 indicadores", libro.hojas[0].filas.length === 6);
const bytes = X.construirLibroXlsx(libro);
t("Excel: ZIP válido (firma PK)", bytes[0] === 0x50 && bytes[1] === 0x4b);
const s2 = Buffer.from(bytes).toString("utf8");
t("Excel: una sola hoja seleccionada", (s2.match(/tabSelected/g) || []).length === 1);
t("Excel: texto escapado", !/<t xml:space="preserve"><img/.test(Buffer.from(X.construirLibroXlsx(L.hojasExcel(s, "x"))).toString("utf8")));

// Fase 2 · Resultados por periodo
const M = (num, den, valor = "", periodo = "") => ({ periodo, num, den, valor, nota: "" });
t("numMiles: 2.950 = 2950", L.numMiles("2.950") === 2950 && L.numMiles("12.345,5") === 12345.5 && L.numMiles("2,5") === 2.5);
// Porcentaje, más es mejor (≥ 95 %)
let e1 = L.evaluarMedicion(ind, M("46", "50"));
t("porcentaje 46/50 = 92 % rojo", cerca(e1.valor, 92) && e1.estado === "rojo");
t("porcentaje 48/50 = 96 % verde", L.evaluarMedicion(ind, M("48", "50")).estado === "verde");
t("porcentaje: 94,96 % redondea a 95,0 → verde", L.evaluarMedicion(ind, M("", "", "94,96")).estado === "verde");
t("porcentaje: valor directo", cerca(L.evaluarMedicion(ind, M("", "", "97,5")).valor, 97.5));
t("porcentaje: num > den = error", !!L.evaluarMedicion(ind, M("51", "50")).error);
t("porcentaje: falta el denominador = error", !!L.evaluarMedicion(ind, M("5", "")).error);
t("porcentaje: texto = error", !!L.evaluarMedicion(ind, M("cinco", "50")).error);
t("medición vacía", L.evaluarMedicion(ind, M("", "")).vacio);
t("textoValor %", L.textoValor(ind, 92) === "92,0 %");
// Menos es mejor (≤ 10 %)
const esp = Object.assign(L.nuevoIndicador("r1", ""), { calculo: "porcentaje", sentido: "menos", estandar: "10" });
t("menos es mejor: 12,6 % rojo", L.evaluarMedicion(esp, M("312", "2.480")).estado === "rojo");
t("menos es mejor: 9,2 % verde", L.evaluarMedicion(esp, M("230", "2.510")).estado === "verde");
// Tasa con base 1.000 y miles en el denominador
const tas = Object.assign(L.nuevoIndicador("r2", ""), { calculo: "tasa", numerador: "caídas", denominador: "estancias", base: "1.000", sentido: "menos", estandar: "1" });
const et = L.evaluarMedicion(tas, M("3", "2.950"));
t("tasa 3/2.950 × 1.000 = 1,02 rojo", cerca(et.valor, 3 / 2950 * 1000) && et.estado === "rojo");
t("textoValor tasa", L.textoValor(tas, 0.645) === "0,65 por 1.000 estancias");
// Centinela
t("centinela 0 verde, 1 rojo", L.evaluarMedicion(cent, M("0", "")).estado === "verde" && L.evaluarMedicion(cent, M("1", "")).estado === "rojo");
t("centinela decimal = error", !!L.evaluarMedicion(cent, M("1,5", "")).error);
// Sin estándar → gris
t("sin estándar = gris", L.evaluarMedicion(Object.assign(L.nuevoIndicador("x", ""), { calculo: "porcentaje" }), M("4", "5")).estado === "gris");
// Media
const med = Object.assign(L.nuevoIndicador("r3", ""), { calculo: "media", sentido: "menos", estandar: "20" });
t("media 18 ≤ 20 verde", L.evaluarMedicion(med, M("", "40", "18")).estado === "verde");
t("media sin valor = error", !!L.evaluarMedicion(med, M("", "40", "")).error);
// LQAS (UPP 95/75 → n 16, c 14)
const upp = Object.assign(L.nuevoIndicador("r4", ""), { calculo: "porcentaje", sentido: "mas", estandar: "95", umbral: "75", objetivo: "decidir", plan: { n: "16", c: "14", nota: "" } });
t("LQAS 15/16 verde", L.evaluarMedicion(upp, M("15", "16")).estado === "verde");
t("LQAS 13/16 rojo", L.evaluarMedicion(upp, M("13", "16")).estado === "rojo");
t("LQAS revisados vacío = n del plan", L.evaluarMedicion(upp, M("14", "")).lqas.rev === 16);
t("LQAS 14 de 14 revisados ya es verde", L.evaluarMedicion(upp, M("14", "14")).estado === "verde");
t("LQAS 7 de 10 revisados ya es rojo (3 fallos > 2)", L.evaluarMedicion(upp, M("7", "10")).estado === "rojo");
t("LQAS 9 de 10 revisados = falta revisar (gris)", L.evaluarMedicion(upp, M("9", "10")).estado === "gris");
t("LQAS más revisados que n = gris", L.evaluarMedicion(upp, M("18", "20")).estado === "gris");
t("LQAS sin plan = error", !!L.evaluarMedicion(Object.assign({}, upp, { plan: { n: "", c: "" } }), M("15", "16")).error);
const uppF = Object.assign({}, upp, { sentido: "menos", estandar: "5", umbral: "25" });
t("LQAS de fallos: 1 fallo de 16 → 15 correctos, verde", L.evaluarMedicion(uppF, M("1", "16")).estado === "verde" && L.evaluarMedicion(uppF, M("1", "16")).lqas.correctos === 15);
t("textoResultado LQAS = casos correctos", L.textoResultado(upp, L.evaluarMedicion(upp, M("13", "16"))) === "13 de 16 correctos");
const urlA = L.enlaceAnalisis(upp, M("13", "16"));
const dA = JSON.parse(decodeURIComponent(urlA.split("?d=")[1]));
t("enlace a «Ya tengo los resultados» (LQAS)", urlA.includes("#/lqas?d=") && dA.an_cum === "13" && dA.an_rev === "16" && dA.pU === "95" && dA.pL === "75");
const conc = Object.assign(L.nuevoIndicador("r5", ""), { calculo: "porcentaje", sentido: "mas", estandar: "90", objetivo: "estimar", muestra: { conf: "95", e: "10", p: "", N: "400" } });
const dPr = JSON.parse(decodeURIComponent(L.enlaceAnalisis(conc, M("64", "78")).split("?d=")[1]));
t("enlace proporción con an_rev/an_cum/an_est", dPr.an_rev === "78" && dPr.an_cum === "64" && dPr.an_est === "90" && dPr.N === "400");
const dPi = JSON.parse(decodeURIComponent(L.enlaceAnalisis(esp, M("8", "80")).split("?d=")[1]));
t("enlace proporción de fallos se invierte", dPi.an_cum === "72" && dPi.an_est === "90");
// Resumen de la serie
tas.resultados = [M("3", "2.950", "", "2026-T1"), M("2", "3.100", "", "2026-T2"), M("1", "3.000", "", "2026-T3")];
const an = L.analizarResultados(tas);
t("serie: 3 válidas, última verde", an.validas.length === 3 && an.ultima.estado === "verde");
t("serie: cumple en 2 de 3 y mejora", /2 de los 3/.test(an.frases[0]) && an.cambio === "mejora");
upp.resultados = [M("15", "16"), M("13", "16"), M("12", "16")];
const an2 = L.analizarResultados(upp);
t("serie LQAS: racha de 2 rojos", an2.frases.some(f => /2 periodos seguidos en rojo/.test(f)) && an2.cambio === "");
// Eje del gráfico
const ej = L.ejeY([92, 96, 98], 95, true);
t("ejeY: incluye valores y estándar, ≤ 100", ej.min <= 92 && ej.max >= 98 && ej.max <= 100 && ej.marcas.length >= 3);
const ej0 = L.ejeY([0, 0], 0, false);
t("ejeY: todo cero no rompe", ej0.max > ej0.min);
// Periodo siguiente
t("siguientePeriodo", L.siguientePeriodo("2026-T4") === "2027-T1" && L.siguientePeriodo("2026-S1") === "2026-S2" && L.siguientePeriodo("2026") === "2027"
  && L.siguientePeriodo("Dic 2026") === "Ene 2027" && L.siguientePeriodo("ene 2026") === "feb 2026" && L.siguientePeriodo("Semana 9") === "Semana 10" && L.siguientePeriodo("") === "");
// Copia de seguridad: se validan los resultados
const sr = L.sanearIndicador({ resultados: [{ periodo: "2026-T1", num: 3, den: "2.950", nota: "x".repeat(900), raro: 1 }, "basura", null] }, "z", "");
t("sanea resultados", sr.resultados.length === 1 && sr.resultados[0].num === "3" && sr.resultados[0].nota.length === 500 && !("raro" in sr.resultados[0]));
t("sanea: sin resultados = []", Array.isArray(L.sanearIndicador({}, "z", "").resultados) && !L.sanearIndicador({}, "z", "").resultados.length);
// Excel con hoja de resultados
const libroR = L.hojasExcel({ indicadores: [tas, upp], compuestos: [] }, "x");
const hR = libroR.hojas.find(h => h.nombre === "Resultados");
t("Excel: hoja Resultados con 6 filas de datos", hR && hR.filas.length === 3 + 6);
t("Excel: semáforo coloreado", hR.filas[3][7].estilo === "rojo" && hR.filas[5][7].estilo === "verde");
const bR = X.construirLibroXlsx(libroR);
t("Excel con resultados: ZIP válido", bR[0] === 0x50 && bR[1] === 0x4b);

// «Traer de Lectura Crítica»: criterio desde la PICO (solo si está vacío) y evidencia en «Observaciones»
{
  const x = L.lecturasDe({ formato: "lecturacritica-lecturas", version: 1, lecturas: [{ id: "l1", titulo: "Valoración precoz del riesgo de úlceras", cita: "Autor C, et al. Int J Nurs Stud. 2022;130:104-12.", tipoNombre: "Estudio de cohortes", decision: "aplicar", pico: { p: "Pacientes ingresados", i: "valoración del riesgo de UPP en las primeras 24 horas", c: "", o: "" }, resultadosClave: [], mensaje: "" }] }).lecturas[0];
  const ind = L.nuevoIndicador("lc1", "");
  t("Lectura Crítica: escribe el criterio si está vacío", L.aplicarLectura(ind, x).criterio && ind.criterio === "En pacientes ingresados, se aplica: valoración del riesgo de UPP en las primeras 24 horas");
  t("Lectura Crítica: evidencia en observaciones", ind.observaciones === "Evidencia: Autor C, et al. Int J Nurs Stud. 2022;130:104-12 (estudio de cohortes). Decisión tras la lectura crítica: aplicar.");
  const ind2 = Object.assign(L.nuevoIndicador("lc2", ""), { criterio: "Mi criterio", observaciones: "Nota previa" });
  t("Lectura Crítica: no toca un criterio escrito", !L.aplicarLectura(ind2, x).criterio && ind2.criterio === "Mi criterio" && ind2.observaciones.startsWith("Nota previa\nEvidencia:"));
  L.aplicarLectura(ind2, x);
  t("Lectura Crítica: no repite la evidencia", ind2.observaciones.split("Evidencia:").length === 2);
  t("Lectura Crítica: pasa el saneado", L.sanearIndicador(ind, "x", "").criterio === ind.criterio);
}

console.log(`${ok} correctos, ${mal} fallidos`);
if (mal) process.exit(1);
