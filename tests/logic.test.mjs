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
t("LQAS URL base", urlL.startsWith("https://cjgaland.github.io/tamano-muestral/#/lqas?d="));
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

console.log(`${ok} correctos, ${mal} fallidos`);
if (mal) process.exit(1);
