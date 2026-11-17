/* Generador del "Programa semanal de mantenimiento" a partir de la plantilla original.
   Funciona igual en el navegador y en Node: generarPrograma(JSZip, plantillaBytes, datos) -> Promise<Uint8Array>
   Copia la plantilla tal cual y solo reescribe la sección 3 (filas de trabajos), los datos de la semana
   y las referencias que dependen del número de filas. */
(function (root) {
  'use strict';
  var COLS = 'ABCDEFGHIJKLMNOPQR'.split('');
  var SH = "'Programa semanal'!";
  // Estructura de la plantilla original (filas)
  var OLD = {
    portales: [
      { fila: 16, dias: [17, 23, 28, 34, 40, 46, 52] },
      { fila: 58, dias: [59, 64, 68, 74, 80, 86, 92] }
    ],
    total: 98
  };
  OLD.portales[0].fines = [22, 27, 33, 39, 45, 51, 57];
  OLD.portales[1].fines = [63, 67, 73, 79, 85, 91, 97];

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function serialFecha(iso) { // "2026-11-23" -> número de serie de Excel
    var p = iso.split('-').map(Number);
    return Math.round((Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  function fraccionHora(h) { // "06:30" -> 0.2708…
    if (!h) return null;
    var p = String(h).split(':').map(Number);
    if (isNaN(p[0])) return null;
    return ((p[0] || 0) * 60 + (p[1] || 0)) / 1440;
  }
  function filasDe(xml) {
    var m, out = {}, re = /<row r="(\d+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
    while ((m = re.exec(xml))) out[+m[1]] = m[0];
    return out;
  }
  function estilos(rowXml) { // columna -> estilo
    var m, out = {}, re = /<c r="([A-Z]+)\d+"(?: s="(\d+)")?/g;
    while ((m = re.exec(rowXml))) out[m[1]] = m[2] || '0';
    return out;
  }
  function celdaOriginal(rowXml, col) {
    var re = new RegExp('<c r="' + col + '\\d+"[^>]*?(?:/>|>[\\s\\S]*?</c>)');
    var m = rowXml.match(re);
    return m ? m[0] : null;
  }
  function atributosFila(rowXml) {
    return rowXml.match(/^<row [^>]*?>/)[0].replace(/^<row r="\d+"/, '').replace(/>$/, '').replace(/\/$/, '');
  }
  // Constructores de celdas
  function cVacia(ref, s) { return '<c r="' + ref + '" s="' + s + '"/>'; }
  function cNum(ref, s, v) { return '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>'; }
  function cTxt(ref, s, v) {
    if (v === undefined || v === null || String(v) === '') return cVacia(ref, s);
    return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
  }
  function cForm(ref, s, f, texto) {
    return '<c r="' + ref + '" s="' + s + '"' + (texto ? ' t="str"' : '') + '><f>' + esc(f) + '</f></c>';
  }
  function cCopia(ref, original) { // copia una celda de la plantilla cambiando su referencia
    return original.replace(/^<c r="[A-Z]+\d+"/, '<c r="' + ref + '"');
  }

  // Inserta o reemplaza una celda en un XML de hoja
  function ponerCelda(xml, ref, celdaXml) {
    var col = ref.match(/^[A-Z]+/)[0], fila = +ref.match(/\d+$/)[0];
    var reFila = new RegExp('<row r="' + fila + '"[^>]*?(?:/>|>[\\s\\S]*?</row>)');
    var m = xml.match(reFila);
    if (!m) { // crear la fila en su lugar
      var filas = filasDe(xml), sig = Object.keys(filas).map(Number).filter(function (r) { return r > fila; }).sort(function (a, b) { return a - b; })[0];
      var nueva = '<row r="' + fila + '">' + celdaXml + '</row>';
      if (sig) return xml.replace(filas[sig], nueva + filas[sig]);
      return xml.replace('</sheetData>', nueva + '</sheetData>').replace('<sheetData/>', '<sheetData>' + nueva + '</sheetData>');
    }
    var fx = m[0];
    if (/\/>$/.test(fx) && !/<c /.test(fx)) fx = fx.replace(/\/>$/, '></row>');
    var reC = new RegExp('<c r="' + ref + '"[^>]*?(?:/>|>[\\s\\S]*?</c>)');
    var nuevaFila;
    if (reC.test(fx)) nuevaFila = fx.replace(reC, celdaXml);
    else {
      var celdas = fx.match(/<c r="[A-Z]+\d+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g) || [];
      var idx = function (c) { var s = c.match(/^[A-Z]+/)[0]; return s.length * 100 + s.charCodeAt(s.length - 1); };
      var despues = celdas.filter(function (c) { return idx(c.match(/<c r="([A-Z]+)/)[1]) > idx(col); })[0];
      nuevaFila = despues ? fx.replace(despues, celdaXml + despues) : fx.replace('</row>', celdaXml + '</row>');
    }
    nuevaFila = nuevaFila.replace(/ spans="[^"]*"/, '');
    return xml.replace(m[0], nuevaFila);
  }
  function estiloDe(xml, ref) {
    var m = xml.match(new RegExp('<c r="' + ref + '"(?: s="(\\d+)")?'));
    return m ? (m[1] || '0') : '0';
  }

  function construirEstructura(datos) {
    // agrupa trabajos por portal y día
    var grupos = [0, 1].map(function () { return [[], [], [], [], [], [], []]; });
    (datos.trabajos || []).forEach(function (t) {
      var p = t.portal === 'Martha' ? 1 : 0;
      var d = Math.max(0, Math.min(6, +t.dia || 0));
      grupos[p][d].push(t);
    });
    grupos.forEach(function (g) {
      g.forEach(function (lista) {
        lista.sort(function (a, b) { return (a.inicio || '99').localeCompare(b.inicio || '99') || (a.orden || 0) - (b.orden || 0); });
      });
    });
    var r = 16, NEW = { portales: [] };
    for (var p = 0; p < 2; p++) {
      var P = { fila: r, dias: [], primeros: [], fines: [], trabajos: grupos[p] };
      r++;
      for (var d = 0; d < 7; d++) {
        P.dias.push(r); r++;
        var n = Math.max(1, grupos[p][d].length);
        P.primeros.push(r); P.fines.push(r + n - 1); r += n;
      }
      NEW.portales.push(P);
    }
    NEW.total = r;
    return NEW;
  }

  // Mapa de filas viejas -> nuevas para reescribir fórmulas de otras hojas
  function crearMapa(NEW) {
    var ini = {}, fin = {}, uno = {};
    ini[OLD.total] = fin[OLD.total] = uno[OLD.total] = NEW.total;
    for (var p = 0; p < 2; p++) {
      var o = OLD.portales[p], n = NEW.portales[p];
      ini[o.fila] = fin[o.fila] = uno[o.fila] = n.fila;
      for (var d = 0; d < 7; d++) {
        ini[o.dias[d]] = fin[o.dias[d]] = uno[o.dias[d]] = n.dias[d];
        ini[o.dias[d] + 1] = n.primeros[d];
        fin[o.fines[d]] = n.fines[d];
      }
    }
    function m(tabla, r) {
      if (r < 16) return r;
      if (tabla[r] === undefined) throw new Error('Referencia no prevista a la fila ' + r);
      return tabla[r];
    }
    return { ini: function (r) { return m(ini, r); }, fin: function (r) { return m(fin, r); }, uno: function (r) { return m(uno, r); } };
  }
  function reescribirRefs(formula, mapa) {
    var re = /'Programa semanal'!(\$?[A-Z]{1,3}\$?)(\d+)(?::(\$?[A-Z]{1,3}\$?)(\d+))?/g;
    return formula.replace(re, function (_, c1, r1, c2, r2) {
      if (c2) return SH + c1 + mapa.ini(+r1) + ':' + c2 + mapa.fin(+r2);
      return SH + c1 + mapa.uno(+r1);
    });
  }
  function reescribirHoja(xml, mapa) {
    return xml.replace(/<f([^>]*)>([^<]*)<\/f>/g, function (todo, attrs, f) {
      if (f.indexOf('Programa semanal') < 0) return todo;
      var crudo = f.replace(/&apos;/g, "'");
      return '<f' + attrs + '>' + reescribirRefs(crudo, mapa) + '</f>';
    });
  }
  function rangos(lista, col) { return lista.map(function (a) { return col + a[0] + ':' + col + a[1]; }).join(' '); }

  function generarPrograma(JSZip, plantilla, datos) {
    return JSZip.loadAsync(plantilla).then(function (zip) {
      var archivos = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml', 'xl/workbook.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels'];
      return Promise.all(archivos.map(function (f) { return zip.file(f).async('string'); })).then(function (t) {
        var s1 = t[0], s2 = t[1], s3 = t[2], s4 = t[3], wb = t[4], ct = t[5], rels = t[6];
        var NEW = construirEstructura(datos);
        var mapa = crearMapa(NEW);
        var filas = filasDe(s1);

        // ---- Hoja Listas: equipos adicionales ----
        var listaA = [], listaF = [], listaG = [];
        for (var i = 4; i <= 33; i++) {
          var fx = filasDe(s4)[i] || '';
          var celdaTexto = function (col) { var c = celdaOriginal(fx, col); return c && /<v>|<t/.test(c); };
          if (celdaTexto('A')) listaA.push(i); if (celdaTexto('F')) listaF.push(i); if (celdaTexto('G')) listaG.push(i);
        }
        var nombresListas = (datos.equiposListas || []).map(function (x) { return x.toLowerCase(); });
        (datos.equiposExtra || []).forEach(function (eq) {
          if (!eq || !eq.nombre) return;
          if (nombresListas.indexOf(eq.nombre.toLowerCase()) >= 0) return;
          var rA = (listaA.length ? listaA[listaA.length - 1] : 3) + 1;
          if (rA > 33) return;
          listaA.push(rA);
          var stA = estiloDe(s4, 'A' + (rA - 1)), stB = estiloDe(s4, 'B' + (rA - 1)), stC = estiloDe(s4, 'C' + (rA - 1)), stD = estiloDe(s4, 'D' + (rA - 1));
          s4 = ponerCelda(s4, 'A' + rA, cTxt('A' + rA, stA, eq.nombre));
          s4 = ponerCelda(s4, 'B' + rA, cTxt('B' + rA, stB, eq.serie || ''));
          s4 = ponerCelda(s4, 'C' + rA, cTxt('C' + rA, stC, eq.portal || 'Pioneer'));
          s4 = ponerCelda(s4, 'D' + rA, cTxt('D' + rA, stD, eq.familia || ''));
          var col = eq.portal === 'Martha' ? 'G' : 'F', lst = eq.portal === 'Martha' ? listaG : listaF;
          var rP = (lst.length ? lst[lst.length - 1] : 3) + 1;
          if (rP <= 33) { lst.push(rP); s4 = ponerCelda(s4, col + rP, cTxt(col + rP, estiloDe(s4, col + (rP - 1)), eq.nombre)); }
          nombresListas.push(eq.nombre.toLowerCase());
        });

        // ---- Hoja Programa semanal: encabezado ----
        var cabecera = [], r;
        for (r = 1; r <= 15; r++) if (filas[r]) cabecera.push(filas[r]);
        var cab = cabecera.join('');
        cab = ponerCelda(cab, 'E7', cNum('E7', estiloDe(cab, 'E7'), serialFecha(datos.lunes)));
        cab = ponerCelda(cab, 'O7', cNum('O7', estiloDe(cab, 'O7'), serialFecha(datos.emision)));
        cab = ponerCelda(cab, 'J7', cTxt('J7', estiloDe(cab, 'J7'), datos.planificador || ''));
        cab = ponerCelda(cab, 'C8', cNum('C8', estiloDe(cab, 'C8'), +datos.tecnicos || 0));
        cab = ponerCelda(cab, 'E8', cNum('E8', estiloDe(cab, 'E8'), +datos.hhTecDia || 0));
        cab = ponerCelda(cab, 'H8', cNum('H8', estiloDe(cab, 'H8'), +datos.diasLab || 0));

        // ---- Sección 3 ----
        var cuerpo = [], numero = 0, filasTrab = [[], []];
        var tFila = filas[18], tVacia = filas[27], sT = estilos(tFila), sV = estilos(tVacia);
        var attrTrab = atributosFila(tFila).replace(/ hidden="1"/, '');
        for (var p = 0; p < 2; p++) {
          var o = OLD.portales[p], n = NEW.portales[p];
          // fila de portal
          var fp = filas[o.fila], st = estilos(fp), celdas = '';
          COLS.forEach(function (c) {
            var ref = c + n.fila;
            if (c === 'F') celdas += cForm(ref, st[c], n.dias.map(function (x) { return 'F' + x; }).join('+'));
            else if (c === 'M') celdas += cForm(ref, st[c], n.dias.map(function (x) { return 'M' + x; }).join('+'));
            else { var orig = celdaOriginal(fp, c); celdas += orig ? cCopia(ref, orig) : cVacia(ref, st[c] || '0'); }
          });
          cuerpo.push('<row r="' + n.fila + '"' + atributosFila(fp) + '>' + celdas + '</row>');
          for (var d = 0; d < 7; d++) {
            var fd = filas[o.dias[d]], sd = estilos(fd), cd = '';
            var a = n.primeros[d], b = n.fines[d];
            COLS.forEach(function (c) {
              var ref = c + n.dias[d];
              if (c === 'D') cd += cForm(ref, sd[c], 'IF($E$7="","",$E$7+' + d + ')');
              else if (c === 'F') cd += cForm(ref, sd[c], 'COUNTA(F' + a + ':F' + b + ')');
              else if (c === 'M') cd += cForm(ref, sd[c], 'SUM(M' + a + ':M' + b + ')');
              else { var orig = celdaOriginal(fd, c); cd += orig ? cCopia(ref, orig) : cVacia(ref, sd[c] || '0'); }
            });
            cuerpo.push('<row r="' + n.dias[d] + '"' + atributosFila(fd).replace(/ hidden="1"/, '') + '>' + cd + '</row>');
            filasTrab[p].push([a, b]);
            var lista = n.trabajos[d];
            for (var k = 0; k < Math.max(1, lista.length); k++) {
              var rr = a + k, t = lista[k], x = '';
              var S = t ? sT : sV;
              var hi = t ? fraccionHora(t.inicio) : null, hf = t ? fraccionHora(t.fin) : null;
              if (t) numero++;
              x += t ? cNum('A' + rr, S.A, numero) : cVacia('A' + rr, S.A);
              x += hi !== null ? cNum('B' + rr, sT.B, hi) : cVacia('B' + rr, S.B);
              x += hf !== null ? cNum('C' + rr, sT.C, hf) : cVacia('C' + rr, S.C);
              x += cTxt('D' + rr, S.D, t && t.equipo);
              x += cForm('E' + rr, S.E, 'IF(D' + rr + '="","",IFERROR(INDEX(Listas!$B$4:$B$33,MATCH(D' + rr + ',Listas!$A$4:$A$33,0)),""))', true);
              x += cTxt('F' + rr, S.F, t && t.ot);
              x += cTxt('G' + rr, S.G, t && t.tipo);
              x += cTxt('H' + rr, S.H, t && t.servicio);
              x += cTxt('I' + rr, S.I, t && t.descripcion);
              x += cTxt('J' + rr, S.J, t && t.prioridad);
              x += t && t.tecnicos !== '' && t.tecnicos != null && !isNaN(+t.tecnicos) ? cNum('K' + rr, S.K, +t.tecnicos) : cVacia('K' + rr, S.K);
              x += cForm('L' + rr, S.L, 'IF(OR(B' + rr + '="",C' + rr + '=""),"",ROUND(MOD(C' + rr + '-B' + rr + ',1)*24,2))');
              x += cForm('M' + rr, S.M, 'IF(OR(L' + rr + '="",K' + rr + '=""),"",L' + rr + '*K' + rr + ')');
              x += cTxt('N' + rr, S.N, t && t.repuestos);
              x += cTxt('O' + rr, S.O, t && t.repListos);
              x += cTxt('P' + rr, S.P, t && t.responsable);
              x += cTxt('Q' + rr, S.Q, t && t.estado);
              x += cTxt('R' + rr, S.R, t && t.obs);
              cuerpo.push('<row r="' + rr + '"' + attrTrab + '>' + x + '</row>');
            }
          }
        }
        // fila total
        var ft = filas[OLD.total], stt = estilos(ft), ctot = '';
        COLS.forEach(function (c) {
          var ref = c + NEW.total;
          if (c === 'M') ctot += cForm(ref, stt[c], 'M' + NEW.portales[0].fila + '+M' + NEW.portales[1].fila);
          else { var orig = celdaOriginal(ft, c); ctot += orig ? cCopia(ref, orig) : cVacia(ref, stt[c] || '0'); }
        });
        cuerpo.push('<row r="' + NEW.total + '"' + atributosFila(ft) + '>' + ctot + '</row>');

        var ini = s1.indexOf('<sheetData>'), fin = s1.indexOf('</sheetData>');
        var cola = s1.slice(fin);
        s1 = s1.slice(0, ini) + '<sheetData>' + cab + cuerpo.join('') + cola;
        s1 = s1.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:R' + NEW.total + '"/>');

        // celdas combinadas
        var merges = (s1.match(/<mergeCell ref="[^"]+"\/>/g) || []).filter(function (m) {
          return +m.match(/[A-Z]+(\d+)/)[1] < 16;
        });
        NEW.portales.forEach(function (P) {
          merges.push('<mergeCell ref="A' + P.fila + ':C' + P.fila + '"/>');
          P.dias.forEach(function (x) { merges.push('<mergeCell ref="A' + x + ':C' + x + '"/>'); });
        });
        merges.push('<mergeCell ref="A' + NEW.total + ':K' + NEW.total + '"/>');
        s1 = s1.replace(/<mergeCells count="\d+">[\s\S]*?<\/mergeCells>/, '<mergeCells count="' + merges.length + '">' + merges.join('') + '</mergeCells>');

        // formato condicional: I12 igual; Q, J, O en todo el bloque de trabajos
        var bloques = s1.match(/<conditionalFormatting sqref="[^"]+">[\s\S]*?<\/conditionalFormatting>/g) || [];
        var nuevosCF = [], vistos = {};
        var ultimo = NEW.total - 1;
        bloques.forEach(function (bq) {
          var sq = bq.match(/sqref="([^"]+)"/)[1];
          if (+sq.match(/\d+/)[0] < 16) { nuevosCF.push(bq); return; }
          var col = sq.match(/^[A-Z]+/)[0];
          if (vistos[col]) return;
          vistos[col] = true;
          var cuerpoCF = bq.replace(/<conditionalFormatting sqref="[^"]+">/, '<conditionalFormatting sqref="' + col + '16:' + col + ultimo + '">');
          nuevosCF.push(cuerpoCF);
        });
        var primerCF = s1.indexOf('<conditionalFormatting'), finCF = s1.lastIndexOf('</conditionalFormatting>') + '</conditionalFormatting>'.length;
        s1 = s1.slice(0, primerCF) + nuevosCF.join('') + s1.slice(finCF);

        // validaciones
        var todos = filasTrab[0].concat(filasTrab[1]);
        s1 = s1.replace(/(<dataValidation type="time"[^>]*?sqref=")[^"]*(")/, function (_, a, b) { return a + rangos(todos, 'B').split(' ').map(function (x) { return x.replace(/^B(\d+):B(\d+)$/, 'B$1:C$2'); }).join(' ') + b; });
        s1 = s1.replace(/(<dataValidation type="whole"[^>]*?sqref=")[^"]*(")/, function (_, a, b) { return a + rangos(todos, 'K') + b; });
        s1 = s1.replace(/<x14:dataValidation [\s\S]*?<\/x14:dataValidation>/g, function (dv) {
          var lista = dv.match(/Listas!\$([A-Z])\$4/)[1];
          var sq;
          if (lista === 'F') sq = rangos(filasTrab[0], 'D');
          else if (lista === 'G') sq = rangos(filasTrab[1], 'D');
          else sq = rangos(todos, { H: 'G', I: 'H', J: 'J', K: 'O', L: 'Q', M: 'P' }[lista]);
          return dv.replace(/<xm:sqref>[^<]*<\/xm:sqref>/, '<xm:sqref>' + sq + '</xm:sqref>');
        });

        // ---- Otras hojas: referencias a Programa semanal ----
        s2 = reescribirHoja(s2, mapa);
        s3 = reescribirHoja(s3, mapa);

        // ---- Libro: área de impresión, recálculo al abrir, sin calcChain ----
        wb = wb.replace(/'Programa semanal'!\$A\$1:\$R\$\d+/, "'Programa semanal'!$A$1:$R$" + (NEW.total + 2));
        wb = wb.replace(/<calcPr[^>]*\/>/, '<calcPr calcId="191029" fullCalcOnLoad="1"/>');
        ct = ct.replace(/<Override PartName="\/xl\/calcChain.xml"[^>]*\/>/, '');
        rels = rels.replace(/<Relationship [^>]*calcChain[^>]*\/>/, '');
        // quitar valores en caché de fórmulas (Excel recalcula al abrir)
        var limpiar = function (x) { return x.replace(/(<\/f>|<f [^>]*\/>)<v>[^<]*<\/v>/g, '$1').replace(/(<\/f>|<f [^>]*\/>)<v\/>/g, '$1'); };

        zip.remove('xl/calcChain.xml');
        zip.file('xl/worksheets/sheet1.xml', limpiar(s1));
        zip.file('xl/worksheets/sheet2.xml', limpiar(s2));
        zip.file('xl/worksheets/sheet3.xml', limpiar(s3));
        zip.file('xl/worksheets/sheet4.xml', s4);
        zip.file('xl/workbook.xml', wb);
        zip.file('[Content_Types].xml', ct);
        zip.file('xl/_rels/workbook.xml.rels', rels);
        return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
      });
    });
  }

  var api = { generarPrograma: generarPrograma, construirEstructura: construirEstructura };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.GeneradorPrograma = api;
})(typeof window !== 'undefined' ? window : this);
