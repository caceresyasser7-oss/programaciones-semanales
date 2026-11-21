# Programa semanal de mantenimiento

Página: https://caceresyasser7-oss.github.io/programaciones-semanales/

- Se agregan los trabajos de la semana con los campos de la fila 15 del formato, y se elige el día de inspección de cada equipo.
- Los datos se guardan en `data/semanas.json` (cada cambio es un commit).
- Cada viernes a las 3:00 PM (Managua) la acción **Corte semanal** cierra la propuesta y deja el Excel en `excel/`, con el formato original (`plantilla.xlsx`).
- Para editar desde un navegador nuevo: botón **Conectar GitHub** y pegar un token con permiso *Contents: Read and write* solo para este repositorio.

Archivos: `index.html` (página), `generador.js` (arma el Excel desde la plantilla), `scripts/corte.js` (corte automático), `.github/workflows/corte.yml`.

## Registro de partes montadas

Página: https://caceresyasser7-oss.github.io/programaciones-semanales/partes.html

- Buscador de números de parte arriba: dice en qué equipos está montado un NP, su estado y cuántas horas le quedan.
- Por equipo: lecturas de horómetros y partes montadas, con el mismo cálculo del Excel *seguimiento de partes* (próximo cambio = horómetro de instalación + vida útil; PRÓXIMO cuando faltan 50 h o menos).
- Los datos se guardan en `data/partes.json` con el mismo token del programa semanal. Botón **Descargar CSV** para abrirlo en Excel.
