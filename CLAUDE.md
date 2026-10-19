# CLAUDE.md — Indicadores

Guía de referencia para el desarrollo asistido por IA. Léela antes de proponer cambios. Responder siempre en español y en lenguaje claro (el autor no es programador).

## 1. Descripción
PWA estática (sin build, sin dependencias, sin CDN) para **crear, definir y comprobar indicadores de calidad y seguridad del paciente**, tal y como se enseña en el curso «Auditorías para la Mejora de la Seguridad del Paciente. Del Mapa de Riesgos a la Mejora Continua» (Pozoblanco, octubre de 2026): tema 3 y ejercicios 5A (caracterización), 5B (definición operativa) y 5C (compuesto). Herramienta docente y de apoyo; no sustituye el criterio profesional.

Hermana de **Tamaño Muestral** (`/Users/Trabajo/Desktop/Tamaño Muestral/`, https://cjgaland.github.io/tamano-muestral/): mismos tokens CSS, modo claro/oscuro, modales (bienvenida, Acerca de, novedades), SW network-first con aviso «Nueva versión disponible», informe imprimible y PDF en iOS.

Autoría (idéntica en el pie de index.html, «Acerca de», el pie del informe en app.js —`informePie()`— y el README; orden alfabético, sin centro de trabajo): Carlos J. Galán Doval (Médico de Familia · Técnico en Calidad y Seguridad del Paciente) y Rafael J. Romero de Castilla Gil (Médico de Familia · Responsable del Área de Calidad y Seguridad del Paciente), Área de Gestión Sanitaria Sur de Córdoba · Andalucía.

## 2. Arquitectura
- `index.html` — esqueleto, header, nav, modales (bienvenida, Acerca de, novedades, confirmar) e `<input type=file>` oculto para importar. Script inline que fija `data-tema` antes de pintar.
- `styles.css` — base copiada de Tamaño Muestral (tokens, modales, «Ver informe») + componentes propios (asistente, tarjetas de opción, calidad, ficha, cuadro, compuesto) + `@media print`. Móvil primero (360 px); 2 columnas ≥ 900 px.
- `logic.js` — funciones PURAS (sin DOM), `window.Logica` / `module.exports`: catálogos (`TIPOS`, `CALCULOS`, `CALIDAD`), modelo (`nuevoIndicador`, `nuevoCompuesto`), `ipr`, `sugerirNombre`, `sentidoSugerido`, `formula`, `unidad`, `textoEstandar`, `avisos` (trampas por paso), `pistasCalidad`, `estadoCalidad`, `acuerdoPiloto`, `pasosCompletos`, `recomendacionMuestra`, `lqasParams`, `enlacesTM`, `calcCompuesto`, `fraseCompuesto`, `sanearIndicador/Compuesto/Datos`, `fusionar`, `hojasExcel`; fase 2: `evaluarMedicion`, `analizarResultados`, `textoValor`/`textoResultado`, `ejeY`, `siguientePeriodo`, `enlaceAnalisis`, `numMiles`.
- `xlsx.js` — Excel real sin dependencias (ZIP STORE + OOXML), adaptado de la skill `app-xlsx-export` en un solo archivo; `window.Xlsx` / `module.exports`. Cambio respecto a la skill: solo la primera hoja lleva `tabSelected` (si todas lo llevan, Excel abre las hojas agrupadas).
- `app.js` — UI: datos, enrutado, asistente, cuadro, compuesto, ficha/informe, aprende, glosario, quiz, PDF iOS, copia/importar/Excel, tema, modales, SW.
- `sw.js` — network-first con caché de respaldo; sin `skipWaiting` al instalar (espera a «Actualizar»).
- `tests/logic.test.mjs` — `node tests/logic.test.mjs` (lógica + Excel).

### Rutas (hash)
`#/` inicio · `#/nuevo` (crea un borrador y redirige) · `#/indicador/<id>/<1-6>` asistente (6 = ficha) · `#/indicador/<id>/resultados` resultados por periodo · `?informe=1` = vista de informe (ficha, `#/cuadro?informe=1`, `#/compuesto/<id>?informe=1`) · `#/cuadro` · `#/compuesto/nuevo` y `#/compuesto/<id>` · `#/ver?d=<JSON>` indicador compartido · `#/aprende` · `#/glosario` · `#/prueba`.

### Datos
- `localStorage["indicadores-datos"]` = `{ app, v, ultimaCopia, indicadores: [], compuestos: [] }`. Se guarda con 250 ms de retardo y al salir (`pagehide`).
- Resultados (fase 2): `ind.resultados = [{ periodo, num, den, valor, nota }]` (máx. 200, en orden de entrada). `valor` solo si no hay num/den (dato ya calculado) o en la media. Recuentos con puntos de millar («2.950») con `numMiles`.
- Al cargar se pasa siempre por `sanearDatos` (claves conocidas, enumerados válidos, textos ≤ 2000).
- **Borradores**: «Nuevo indicador» / «Nuevo compuesto» crean el objeto en memoria (`borradores`) y solo se guarda en `datos` cuando se escribe algo (`esVacio`, `esVacioCmp`).
- Campos con `data-k="ruta.anidada"` (p. ej. `muestra.e`, `calidad.E.t`, `rows.0.cum`) → `setPath`. Botones `data-pick` (excluyentes, repintan el paso conservando el scroll), `data-set` (chips rápidos), `data-multi` (fuentes). Al escribir, `refrescarZonas()` actualiza solo `[data-zona]` (IPR, avisos, fórmula, estándar, resumen, semáforo, piloto) y los `href` de `[data-zona-href]`, sin perder el foco.
- Copia de seguridad: descarga `.json` (en iOS, menú Compartir) y guarda `ultimaCopia`; aviso en el cuadro si no hay copia o tiene más de 30 días. Importar: añadir (fusiona por id, gana el más reciente) o sustituir (con doble confirmación).
- Compartir: `#/ver?d=` con el indicador compacto (sin id ni campos vacíos); se valida con `sanearIndicador` y todo se pinta con `esc()` (probado con `<img onerror>`: no se ejecuta).

## 3. Contenido y decisiones didácticas
- Asistente: 1 Del riesgo al indicador (subproceso, hallazgo, G·O·D → IPR = G×O×D, criterio, nombre con «Proponer nombre») · 2 Caracterizarlo (5A: qué mide + justificación; forma de cálculo con «¿con muestra?») · 3 Definición operativa (5B, campos según la forma de cálculo) · 4 Calidad (SMART + válido, fiable, útil, comparable: Sí/Dudoso/No + «cómo se cumple», pistas automáticas, semáforo; pilotaje con tabla 2×2 y % de acuerdo) · 5 Muestra (estimar / decidir LQAS / media / compuesto / todos los casos; enlace a Tamaño Muestral; se anotan n y c) · 6 Ficha.
- **Estándar y umbral**: «Estándar o meta» = el «umbral / límite» de la plantilla 5B; «Umbral de alarma» (opcional) solo para LQAS. Decisión aprobada por el autor (2026-10-19).
- **Sentido** (más es mejor ≥ / menos es mejor ≤): `sentidoSugerido` detecta por palabras (errores, discrepancias, caídas, demoras, «más de N», «sin registro»…; quita antes «riesgo de…», «escala de…», «prevención de…») y `avisos` marca la trampa del sentido invertido. Motivo: en el Anexo I de la auditoría modelo (`00 Auditoría Interna.pdf`) los indicadores 1 («% de discrepancias…») y 3 («% que superan el tiempo… sin reevaluación») llevan ≥ 95 % / ≥ 90 % (deberían ir con ≤). Se avisó al autor.
- SMART con **Relevante** (tema 3 y diapositivas). La plantilla del 5A pone «Realista»; se explica en «Explicado fácil».
- **Tasa** según el tema 3 (casos referidos a una población y un periodo, p. ej. por 1.000 estancias), no como «porcentaje referido a un periodo» (unidad temática de Rafa).
- Centinela: al elegirlo se fija sentido «menos», estándar 0 y objetivo «sin muestra».
- LQAS con indicadores de fallos: `lqasParams` convierte al complementario (≤ 5 % fallos → pU 95; umbral 20 % fallos → pL 80) y se explica en el paso 5.
- **Esta app no calcula muestras ni LQAS ni kappa**: enlaza con Tamaño Muestral (`enlacesTM`). Valores como texto y en %: `proporcion` (conf, e, p [50 si vacío], N, loss), `lqas` (pU, pL, alpha = falsa alarma, beta = fuego sin alarma, N, loss), `media` (conf, sigma, e, N, loss), `compuesto` (pac, rows), `concordancia` (a, b, c, d). Comprobado en la web real: 95/75 α5 β20 → 16/14.
- El compuesto sí se calcula aquí (es una suma): Σ cumplimientos ÷ Σ oportunidades, % por criterio, peso, punto débil, media de % solo para comparar, frase «titular + desglose» (editable). Avisos: el titular esconde al flojo (> 10 puntos por debajo), oportunidades ≠ pacientes, media ≠ compuesto.
- Ejemplos SOLO de la parte teórica o inventados (NO de los ejercicios prácticos, que estropearían el solucionario): UPP 24 h (LQAS 95/75 → 16/14), conciliación al ingreso y eventos adversos (minicaso del tema 3; estimar N 400 ±10 % → 78), Ottawa (90/75 → 40/33), registro del dolor (85/60 → 22/16), cirugía en lugar equivocado (centinela), caídas por 1.000 estancias (tasa), esperas > 30 min (de media a %), dispensadores de solución hidroalcohólica (estructura), compuesto del infarto (28/40 = 70 %; media 63,9 %). Planes calculados con `stats.js` de Tamaño Muestral.
- Evitar en ejemplos y quiz: triaje/constantes, sala de espera/reevaluación, pulsera en traslados, checklist quirúrgico, EPOC, valoración de caídas en ≥ 65 años, conciliación en urgencias (son de los ejercicios).
- Excel: hojas «5A Caracterización», «5B Definición operativa», si hay, «5C Compuestos» (punto débil resaltado) y «Resultados» (semáforo coloreado).
- **Resultados (fase 2, Verificar del PDCA)**, aprobado por el autor el 2026-10-19: pantalla aparte (desde la ficha y la tarjeta del cuadro), una fila por periodo con semáforo verde/rojo según el sentido (se compara el valor redondeado a los decimales que se muestran: 0 en centinela/absoluto, 2 en tasa/ratio, 1 en el resto); gris sin estándar. Centinela: cualquier caso = rojo. **LQAS**: solo se aplica la regla n/c del paso 5 (correctos = n − fallos si «menos es mejor»; decide antes de terminar como Tamaño Muestral) y se muestra «x de n correctos», nunca un % como dato; tira de semáforos en vez de línea. Con muestra (estimar): aviso de margen de error y enlace a «Ya tengo los resultados» de Tamaño Muestral (`an_rev`, `an_cum`, `an_est`, invertidos si es de fallos). Tendencia sin estadística: cumple en X de los últimos 4, racha de rojos, mejora/empeora respecto al anterior, aviso con < 3 periodos. Gráfico SVG propio (viewBox 400, colores por CSS para el modo oscuro). Ejemplos con mediciones inventadas.
- Informe/impresión: «Ver informe» añade `?informe=1` (clase `body.ver-informe`); «Guardar / imprimir» en escritorio/Android va al informe y llama a `window.print()`; en iOS genera el PDF (`generarPDF` + `pdfDesdeJPEGs`, copiados de Tamaño Muestral) y abre Compartir.
- Los `textarea.txt-in` crecen solos (`autoAlto` + MutationObserver).
- «Explicado fácil» recuerda si se dejó abierto (`indicadores-facil-abierto`).

## 4. Entorno y pruebas
- Node v26 (Homebrew). Si falla por `libsimdjson`: `brew upgrade node`. R 4.6.1 (`Rscript`).
- Servir en local: `python3 -m http.server 8766` desde la carpeta (Tamaño Muestral usa el 8765). El navegador integrado de Claude Code no admite service workers (el error «fetching the script» en consola es eso).
- Revisar siempre en escritorio y en móvil (360 px): sin scroll horizontal (`document.documentElement.scrollWidth === 360`).

## 5. Versiones y actualización
- `APP_VERSION` y `NOVEDADES` al principio de `app.js` (formato AAAA.NN, la más reciente arriba). Claves de localStorage: `indicadores-version-vista`, `indicadores-bienvenida-vista`, `indicadores-tema`, `indicadores-facil-abierto`, `indicadores-datos`.

## 6. Copias de seguridad locales
Carpeta `Backup/` (ignorada en git) con las 5 copias más recientes:
`bash ~/.claude/skills/app-copia_seguridad/scripts/copia_seguridad.sh --root "/Users/Trabajo/Desktop/Indicadores" --name Indicadores`. Hacer una antes de cada publicación.

## 7. Publicación
- Repositorio público `cjgaland/indicadores-calidad`; GitHub Pages desde `main` (raíz): https://cjgaland.github.io/indicadores-calidad/. Creado y publicado con permiso del autor el 2026-10-19 (v2026.01; Pages tarda ~40 s). Cada nueva publicación (push) se hace solo cuando el autor lo pida.
- Pasos para publicar una versión: (1) copia de seguridad; (2) nueva entrada en `NOVEDADES` + `APP_VERSION`; (3) mismos `?v=` en index.html (styles, logic, xlsx, app); (4) subir `CACHE_NAME` en sw.js; (5) `node tests/logic.test.mjs`; (6) commit en español (terminado en `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`) y `git push`.
- `PROMPT_NUEVA_SESION.md` no se publica (está en `.gitignore`).

## 8. Pendiente
- Fase 2 hecha en local (2026-10-19), pendiente de publicar: al publicar, entrada en `NOVEDADES` y `APP_VERSION` 2026.02.
- Probar en un iPhone real el PDF y la copia de seguridad por Compartir, y la instalación como PWA.
- Varios cuadros con nombre (la estructura de datos lo permitiría más adelante).
- Manual de usuario (skill `app-user-manual`) cuando la app esté estable.
