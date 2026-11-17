// Corte semanal: cierra las semanas vencidas (viernes 15:00 Managua) y genera su Excel en excel/.
// Lo ejecuta GitHub Actions cada viernes; también se puede correr a mano: node scripts/corte.js
const fs = require('fs'), path = require('path');
const raiz = path.join(__dirname, '..');
const JSZip = require(path.join(raiz, 'vendor/jszip.min.js'));
const G = require(path.join(raiz, 'generador.js'));

const MARTHA = ['Pala ST1030 #1','Pala ST1030 #2','Pala ST1030 #3','Pala ST1030 #6','Simba S7D #1','Simba S7D #3','Boomer 282 #1','Boomer 282 #2','Boltec S10 S #2'];
const CATALOGO = ['Pala ST1030 #4','Pala ST1030 #5','Simba S7D #2','Boomer 282 #3','Boltec S #1','ST3.5 #6','Wagner #5','Scalemin #2','Elevador S44 #1'].concat(MARTHA, ['Scalemin #1']);
const DEFAULT_CFG = { planificador: 'Yasser Caceres Rios', tecnicos: 11, hhTecDia: 12, diasLab: 7 };

const iso = (y, m, d) => new Date(Date.UTC(y, m, d)).toISOString().slice(0, 10);
const addD = (s, n) => { const p = s.split('-').map(Number); return iso(p[0], p[1] - 1, p[2] + n); };
const corteDe = (lunes) => { const p = lunes.split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2] - 3, 21, 0, 0); };
function semanaISO(s) { const p = s.split('-').map(Number); const d = new Date(Date.UTC(p[0], p[1] - 1, p[2])); const wd = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - wd + 3); const f = new Date(Date.UTC(d.getUTCFullYear(), 0, 4)); return 1 + Math.round(((d - f) / 86400000 - 3 + ((f.getUTCDay() + 6) % 7)) / 7); }
// lunes de la semana cuyo corte fue el viernes más reciente (ya pasado)
function lunesUltimoCorte(ahora) {
  const m = new Date(ahora - 6 * 3600 * 1000); // hora de Managua
  const wd = (m.getUTCDay() + 6) % 7; // lunes=0
  let atras = (wd - 4 + 7) % 7; if (wd === 4 && m.getUTCHours() < 15) atras = 7;
  return addD(iso(m.getUTCFullYear(), m.getUTCMonth(), m.getUTCDate() - atras), 3);
}

const archivo = path.join(raiz, 'data/semanas.json');
const semanas = fs.existsSync(archivo) ? JSON.parse(fs.readFileSync(archivo, 'utf8') || '{}') : {};
const ahora = Date.now();
const objetivo = lunesUltimoCorte(ahora);
let cambios = false;

// si nadie abrió la página esa semana, igual se genera el programa (vacío) del corte de hoy
if (!semanas[objetivo] && ahora - corteDe(objetivo) < 24 * 3600 * 1000) {
  const previas = Object.keys(semanas).sort(); const ult = previas.length ? semanas[previas[previas.length - 1]] : null;
  semanas[objetivo] = { lunes: objetivo, semana: semanaISO(objetivo), estado: 'abierta', creadaEn: new Date().toISOString(), config: (ult && ult.config) || DEFAULT_CFG, diasEquipo: {}, equiposExtra: (ult && ult.equiposExtra) || [], trabajos: [] };
}
for (const l of Object.keys(semanas)) {
  const s = semanas[l];
  if (s.estado === 'abierta' && ahora >= corteDe(l)) {
    Object.assign(s, { estado: 'cerrada', emision: addD(l, -3), cerradaEn: new Date().toISOString(), cerradaPor: 'automático' });
    cambios = true; console.log('Semana cerrada:', l);
  }
}
if (cambios) fs.writeFileSync(archivo, JSON.stringify(semanas, null, 1));

const nombre = (s) => `Programa_Semanal_Mantenimiento_S${s.semana}_${s.lunes}.xlsx`;
const plantilla = fs.readFileSync(path.join(raiz, 'plantilla.xlsx'));
const tareas = Object.values(semanas).filter((s) => s.estado === 'cerrada' && !fs.existsSync(path.join(raiz, 'excel', nombre(s))));
(async () => {
  for (const s of tareas) {
    const cfg = s.config || DEFAULT_CFG, extras = s.equiposExtra || [];
    const datos = { lunes: s.lunes, emision: s.emision || addD(s.lunes, -3), planificador: cfg.planificador, tecnicos: cfg.tecnicos, hhTecDia: cfg.hhTecDia, diasLab: cfg.diasLab,
      equiposListas: CATALOGO, equiposExtra: extras,
      trabajos: (s.trabajos || []).map((t, i) => { const e = extras.find((x) => x.nombre === t.equipo); return Object.assign({}, t, { orden: i, portal: t.portal || ((MARTHA.includes(t.equipo) || (e && e.portal === 'Martha')) ? 'Martha' : 'Pioneer') }); }) };
    const bytes = await G.generarPrograma(JSZip, plantilla, datos);
    fs.writeFileSync(path.join(raiz, 'excel', nombre(s)), Buffer.from(bytes));
    console.log('Excel generado:', nombre(s));
  }
})();
