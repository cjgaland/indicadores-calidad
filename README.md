# Indicadores

PWA estática (HTML/CSS/JS puro, sin dependencias ni CDN, funciona sin conexión) para crear, definir y comprobar indicadores de calidad y seguridad del paciente, paso a paso, como en el curso «Auditorías para la Mejora de la Seguridad del Paciente. Del Mapa de Riesgos a la Mejora Continua».

Asistente en seis pasos: del hallazgo del Mapa de Riesgos al indicador, caracterización (tipo y forma de cálculo), definición operativa, comprobación de su calidad (SMART, válido, fiable, útil, con pilotaje), muestra (enlazando con [Tamaño Muestral](https://cjgaland.github.io/tamano-muestral/)) y ficha imprimible. Además: indicador compuesto con su desglose, cuadro de indicadores exportable a Excel, copia de seguridad e importación, enlace para compartir un indicador, «Aprende», glosario y «Ponte a prueba». Los datos se guardan solo en el dispositivo.

## Uso
Servir la carpeta y abrir `index.html`:

    python3 -m http.server 8766

Tests de la lógica y del Excel:

    node tests/logic.test.mjs

Documentación técnica y decisiones: ver `CLAUDE.md`.

Aplicación desarrollada por:

- **Carlos J. Galán Doval** · Médico de Familia · Técnico en Calidad y Seguridad del Paciente
- **Rafael J. Romero de Castilla Gil** · Médico de Familia · Responsable del Área de Calidad y Seguridad del Paciente

Área de Gestión Sanitaria Sur de Córdoba · Andalucía. De libre uso con fines docentes y divulgativos. No sustituye el criterio profesional.
