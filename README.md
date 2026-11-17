# Programa semanal de mantenimiento

Página: https://caceresyasser7-oss.github.io/programaciones-semanales/

- Se agregan los trabajos de la semana con los campos de la fila 15 del formato, y se elige el día de inspección de cada equipo.
- Los datos se guardan en `data/semanas.json` (cada cambio es un commit).
- Cada viernes a las 3:00 PM (Managua) la acción **Corte semanal** cierra la propuesta y deja el Excel en `excel/`, con el formato original (`plantilla.xlsx`).
- Para editar desde un navegador nuevo: botón **Conectar GitHub** y pegar un token con permiso *Contents: Read and write* solo para este repositorio.

Archivos: `index.html` (página), `generador.js` (arma el Excel desde la plantilla), `scripts/corte.js` (corte automático), `.github/workflows/corte.yml`.
