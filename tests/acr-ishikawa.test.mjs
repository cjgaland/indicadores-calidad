// «Traer de ACR · Ishikawa»: el lector y la creación con el paquete REAL que genera la app ACR · Ishikawa
// a partir de su ejemplo de la sepsis (si su carpeta está al lado en el Escritorio). node tests/acr-ishikawa.test.mjs
import { createRequire } from "module";
import { existsSync } from "fs";
import { fileURLToPath } from "url";
const require = createRequire(import.meta.url);
const L = require("../logic.js");
let ok = 0, mal = 0;
const prueba = (nombre, fn) => { try { fn(); ok++; } catch (e) { mal++; console.log("✗ " + nombre + "\n   " + e.message); } };
const igual = (a, b, m = "") => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${m} esperaba ${JSON.stringify(b)} y es ${JSON.stringify(a)}`); };
const si = (v, m = "") => { if (!v) throw new Error(m || "esperaba verdadero"); };
let n = 0; const genId = () => "t" + (++n);
const ESTA_APP = L.auditoriaDesdeAcr ? "auditorias" : "indicadores";

prueba("formato y versión", () => {
  igual(L.analisisDeAcr(null).error, "formato"); igual(L.analisisDeAcr({ formato: "otro" }).error, "formato");
  igual(L.analisisDeAcr({ formato: "acrishikawa-analisis", version: 2, analisis: [] }).error, "version");
  igual(L.analisisDeAcr({ formato: "acrishikawa-analisis", version: 1, analisis: [] }).analisis, []);
});
prueba("estándar desde el nivel esperado", () => {
  igual(L.estandarDeEsperado("≥ 90 %"), { sentido: "mas", estandar: "90" }); igual(L.estandarDeEsperado("≤ 2,5 %"), { sentido: "menos", estandar: "2,5" }); igual(L.estandarDeEsperado("Sí").estandar, "");
});
prueba("saneado: fuera ejemplos, acciones vacías y campos raros", () => {
  const r = L.analisisDeAcr({ intercambio: { formato: "acrishikawa-analisis", version: 1, analisis: [{ id: "ej-x", acciones: [{ que: "a" }] }, { id: "a1", tipo: "raro", sac: 9, acciones: [{ que: "", fuerza: "x" }, { id: "x1", que: "Hacer algo", fuerza: "magia", causas: ["c", 3], indicador: { nombre: "" } }] }] } });
  igual(r.analisis.length, 1); igual(r.analisis[0].tipo, "ishikawa"); igual(r.analisis[0].sac, null);
  igual(r.analisis[0].acciones.length, 1); igual(r.analisis[0].acciones[0].fuerza, ""); igual(r.analisis[0].acciones[0].indicador, null); igual(r.analisis[0].acciones[0].causas, ["c", "3"]);
});

const DIR = fileURLToPath(new URL("../../../ACR Ishikawa/Aplicación/", import.meta.url));
if (!existsSync(DIR + "logic.js")) console.log("· (sin la carpeta de ACR · Ishikawa: se saltan las pruebas con su paquete real)");
else {
  const LA = require(DIR + "logic.js"), CA = require(DIR + "contenido.js");
  const ej = JSON.parse(JSON.stringify(CA.EJEMPLOS.analisis.find(a => a.id === "ej-acr-sepsis")).replace(/"ej-/g, '"x-'));
  const paquete = LA.paqueteIntercambio([LA.sanearAnalisis(ej, genId, "2027-02-05")], "x", "2027-02-05T10:00:00Z");
  const r = L.analisisDeAcr(paquete), an = r.analisis[0];
  prueba("paquete real: el ACR de la sepsis con sus tres acciones", () => {
    si(!r.error); igual(r.analisis.length, 1); igual(an.tipo, "acr"); igual(an.sac, 1); igual(an.cerrado, true);
    igual(an.acciones.map(x => x.fuerza), ["fuerte", "intermedia", "intermedia"]); si(an.acciones.every(x => x.indicador && x.causas.length));
    igual(L.analisisDeAcr({ app: "acrishikawa", intercambio: paquete }).analisis.length, 1);
  });
  if (ESTA_APP === "auditorias") prueba("Auditorías: una auditoría nueva con un criterio por acción y su indicador", () => {
    const a = L.auditoriaDesdeAcr(an, genId, "2027-02-10T10:00:00Z", [an.acciones[0].id, an.acciones[1].id]);
    igual(a.origen.tipo, "acr"); si(/ACR · Ishikawa/.test(a.origen.detalle)); si(/SAC 1/.test(a.origen.detalle));
    igual(a.criterios.length, 2); igual(a.criterios[0].criterio, an.acciones[0].que); si(/Acción fuerte/.test(a.criterios[0].hallazgo));
    igual(a.criterios[0].ind.estandar, "95"); igual(a.criterios[0].ind.sentido, "mas"); igual(a.criterios[1].ind.tipo, "resultado");
    igual(L.auditoriaDesdeAcr(an, genId, "2027-02-10T10:00:00Z").criterios.length, 3);
    si(L.ORIGENES.acr);
  });
  if (ESTA_APP === "indicadores") prueba("Indicadores: un indicador nuevo desde la acción", () => {
    const x = an.acciones[1], i = L.indicadorDesdeAccion(an, x, "i1", "2027-02-10T10:00:00Z");
    igual(i.nombre, x.indicador.nombre); igual(i.criterio, x.que); igual([i.estandar, i.sentido, i.tipo, i.calculo], ["90", "mas", "", "porcentaje"]);
    igual(i.numerador, x.indicador.numerador); igual(i.responsable, x.responsable); si(/ACR · Ishikawa/.test(i.observaciones)); si(/intermedia/.test(i.hallazgo));
    const v = L.validarIndicador ? L.validarIndicador(i) : null; if (v) si(v);
  });
}
console.log(`\n${ok} pruebas bien, ${mal} mal`);
if (mal) process.exit(1);
