/* =====================================================================
   Aria Escalón · Generador del Reporte Comercial (pdfmake docDefinition)
   ---------------------------------------------------------------------
   Toda la lógica de diseño del PDF vive aquí. Recibe registros ya
   normalizados (modelo, unidad, nivel, mts2, precio) y devuelve el
   docDefinition de pdfmake. También expone utilidades de normalización
   y resumen. Funciona en navegador (window.AriaReport) y en Node
   (module.exports) para poder testear el diseño sin abrir el navegador.
   ===================================================================== */
(function (root) {
  'use strict';

  // ---- Paleta de marca (cálida / cobre) -----------------------------
  var PALETTE = {
    paper:     '#f6f3ee', // fondo de página (crema cálido)
    ink:       '#2c2926', // texto principal
    inkMid:    '#5f5950', // texto secundario
    inkSoft:   '#938b80', // etiquetas / muted
    line:      '#d6cfc3', // hairline
    lineSoft:  '#e6e0d6', // separador muy suave
    panel:     '#efeae1', // fondo barra de estadísticas
    copper:    '#a96a3b', // acento (letras de modelo, valores)
    copperDeep:'#8c5630'
  };

  var MESES = ['enero','febrero','marzo','abril','mayo','junio','julio',
               'agosto','septiembre','octubre','noviembre','diciembre'];

  // ---- Formateadores ------------------------------------------------
  function formatMoney(n) {
    var v = Number(n) || 0;
    return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function formatArea(n) {
    var v = Number(n) || 0;
    return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function formatInt(n) {
    return (Number(n) || 0).toLocaleString('en-US');
  }
  function todayEs(d) {
    d = d || new Date();
    return d.getDate() + ' de ' + MESES[d.getMonth()] + ', ' + d.getFullYear();
  }

  // ---- Normalización de filas crudas del Excel ----------------------
  // Acepta encabezados en español con o sin acentos / variaciones.
  function stripAccents(s) {
    return String(s == null ? '' : s)
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .toLowerCase().trim();
  }

  var HEADER_ALIASES = {
    modelo: ['modelo', 'model', 'tipologia', 'tipo'],
    unidad: ['unidad', 'numero de unidad', 'no de unidad', 'no unidad', 'num unidad',
             'numero unidad', 'apto', 'apartamento', 'unit', 'no', '#', 'numero'],
    nivel:  ['nivel', 'piso', 'level', 'floor', 'planta'],
    mts2:   ['mts2 totales', 'metros totales', 'metros cuadrados totales', 'm2 totales',
             'area total', 'mts2', 'm2', 'metros', 'area', 'superficie', 'mts totales',
             'mts² totales', 'm² totales'],
    precio: ['precio con bodega', 'precio total', 'precio', 'price', 'valor', 'monto',
             'precio bodega', 'precio + bodega', 'precio mas bodega']
  };

  function resolveHeaderMap(headerRow) {
    // headerRow: array de nombres de columna del Excel
    var map = {}; // campo -> índice
    var normalized = headerRow.map(stripAccents);
    Object.keys(HEADER_ALIASES).forEach(function (field) {
      var aliases = HEADER_ALIASES[field];
      var idx = -1;
      // 1) match exacto
      for (var a = 0; a < aliases.length && idx === -1; a++) {
        idx = normalized.indexOf(aliases[a]);
      }
      // 2) match por "contiene"
      if (idx === -1) {
        for (var i = 0; i < normalized.length && idx === -1; i++) {
          for (var j = 0; j < aliases.length; j++) {
            if (normalized[i] && normalized[i].indexOf(aliases[j]) !== -1) { idx = i; break; }
          }
        }
      }
      map[field] = idx;
    });
    return map;
  }

  function toNumber(v) {
    if (typeof v === 'number') return v;
    if (v == null) return NaN;
    var s = String(v).replace(/[^0-9.,-]/g, '').trim();
    if (s === '') return NaN;
    // Si hay coma y punto, asumimos coma = miles
    if (s.indexOf(',') !== -1 && s.indexOf('.') !== -1) {
      s = s.replace(/,/g, '');
    } else if (s.indexOf(',') !== -1 && s.indexOf('.') === -1) {
      // coma como decimal (es-SV poco común en estos exports, pero por si acaso)
      var parts = s.split(',');
      if (parts[parts.length - 1].length === 2) s = s.replace(/,/g, '.');
      else s = s.replace(/,/g, '');
    }
    return parseFloat(s);
  }

  // rows: array de arrays (AOA), primera fila = encabezados
  function normalizeRows(rows) {
    if (!rows || !rows.length) return { records: [], errors: ['El archivo está vacío.'] };

    // Buscar la fila de encabezados (la primera con texto significativo)
    var headerIdx = 0;
    for (var r = 0; r < Math.min(rows.length, 10); r++) {
      var hasText = rows[r].some(function (c) { return typeof c === 'string' && c.trim().length > 1; });
      if (hasText) { headerIdx = r; break; }
    }
    var header = rows[headerIdx].map(function (c) { return c == null ? '' : c; });
    var map = resolveHeaderMap(header);

    var missing = [];
    ['modelo', 'unidad', 'nivel', 'mts2', 'precio'].forEach(function (f) {
      if (map[f] === -1) missing.push(f);
    });
    if (missing.length) {
      return {
        records: [],
        errors: ['No se encontraron las columnas: ' + missing.join(', ') +
                 '. Encabezados detectados: ' + header.filter(Boolean).join(' | ')]
      };
    }

    var records = [];
    var warnings = [];
    for (var i = headerIdx + 1; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.every(function (c) { return c == null || String(c).trim() === ''; })) continue;
      var modelo = String(row[map.modelo] == null ? '' : row[map.modelo]).trim();
      var unidadRaw = row[map.unidad];
      var nivelRaw = row[map.nivel];
      var mts2 = toNumber(row[map.mts2]);
      var precio = toNumber(row[map.precio]);
      if (!modelo) continue;
      records.push({
        modelo: modelo.toUpperCase().replace(/^MODELO\s+/i, '').trim(),
        unidad: String(unidadRaw == null ? '' : unidadRaw).trim(),
        nivel: String(nivelRaw == null ? '' : nivelRaw).trim(),
        nivelNum: toNumber(nivelRaw),
        unidadNum: toNumber(unidadRaw),
        mts2: isNaN(mts2) ? 0 : mts2,
        precio: isNaN(precio) ? 0 : precio
      });
    }
    return { records: records, errors: [], warnings: warnings };
  }

  // ---- Agrupación y resumen ----------------------------------------
  function groupByModel(records) {
    var groups = {};
    records.forEach(function (rec) {
      (groups[rec.modelo] = groups[rec.modelo] || []).push(rec);
    });
    var keys = Object.keys(groups).sort(function (a, b) {
      return a.localeCompare(b, 'es', { numeric: true });
    });
    return keys.map(function (k) {
      var list = groups[k].slice().sort(function (a, b) {
        if (a.nivelNum !== b.nivelNum) return (a.nivelNum || 0) - (b.nivelNum || 0);
        return (a.unidadNum || 0) - (b.unidadNum || 0);
      });
      return { modelo: k, units: list };
    });
  }

  function buildSummary(records) {
    var precios = records.map(function (r) { return r.precio; }).filter(function (p) { return p > 0; });
    var niveles = {}; records.forEach(function (r) { if (r.nivel) niveles[r.nivel] = true; });
    var modelos = {}; records.forEach(function (r) { modelos[r.modelo] = true; });
    var modelKeys = Object.keys(modelos).sort(function (a, b) { return a.localeCompare(b, 'es', { numeric: true }); });
    return {
      total: records.length,
      niveles: Object.keys(niveles).length,
      modelos: modelKeys.length,
      modelosList: modelKeys,
      precioMin: precios.length ? Math.min.apply(null, precios) : 0,
      precioMax: precios.length ? Math.max.apply(null, precios) : 0
    };
  }

  // ---- Etiquetas de fila -------------------------------------------
  function unitLabel(rec) {
    var n = rec.nivel || (isNaN(rec.nivelNum) ? '' : rec.nivelNum);
    var u = rec.unidad || (isNaN(rec.unidadNum) ? '' : rec.unidadNum);
    if (n !== '' && u !== '') return 'APTO ' + n + '-' + u;
    if (u !== '') return 'APTO ' + u;
    return 'APTO';
  }
  function levelLabel(rec) {
    var n = rec.nivel || (isNaN(rec.nivelNum) ? '' : rec.nivelNum);
    return n !== '' ? 'Nivel ' + n : '—';
  }

  // ---- Texto serif con respaldo para glifos faltantes --------------
  // La versión TRIAL de Albra no incluye '+' (ni algunos símbolos), así
  // que esos caracteres se dibujan con PP Mori para evitar el "tofu" (□).
  var ALBRA_UNSAFE = /[+#&@*]/;
  function serifRuns(str, size, color) {
    var runs = [];
    var buffer = '';
    var bufferSafe = null;
    function flush() {
      if (buffer === '') return;
      if (bufferSafe) {
        runs.push({ text: buffer, font: 'Albra', fontSize: size, color: color });
      } else {
        runs.push({ text: buffer, font: 'PPMori', fontSize: size * 0.82, color: color });
      }
      buffer = '';
    }
    for (var i = 0; i < str.length; i++) {
      var ch = str[i];
      var safe = !ALBRA_UNSAFE.test(ch);
      if (bufferSafe === null) bufferSafe = safe;
      if (safe !== bufferSafe) { flush(); bufferSafe = safe; }
      buffer += ch;
    }
    flush();
    return runs;
  }

  // ---- Componentes del documento -----------------------------------
  function statCell(label, value, sub) {
    return {
      stack: [
        { text: label, font: 'PPMori', fontSize: 6.5, characterSpacing: 1.1, color: PALETTE.inkSoft, margin: [0, 0, 0, 7] },
        { text: value, font: 'Albra', fontSize: 18, color: PALETTE.ink, margin: [0, 0, 0, 3], lineHeight: 1 },
        { text: sub, font: 'PPMori', fontSize: 7, color: PALETTE.inkSoft }
      ]
    };
  }

  function statBar(summary) {
    return {
      table: {
        widths: ['*', '*', '*', '*'],
        body: [[
          statCell('APARTAMENTOS DISPONIBLES', formatInt(summary.total),
                   'En ' + summary.niveles + ' niveles activos'),
          statCell('MODELOS ACTIVOS', formatInt(summary.modelos),
                   summary.modelosList.join('  ·  ')),
          { stack: [
              { text: 'PRECIO MÍNIMO', font: 'PPMori', fontSize: 6.5, characterSpacing: 1.1, color: PALETTE.inkSoft, margin: [0, 0, 0, 7] },
              { text: [{ text: 'US$ ', font: 'PPMori', fontSize: 9, color: PALETTE.inkSoft }, { text: formatMoney(summary.precioMin), font: 'Albra', fontSize: 16, color: PALETTE.copper }], margin: [0, 0, 0, 3] },
              { text: 'Con bodega', font: 'PPMori', fontSize: 7, color: PALETTE.inkSoft }
          ] },
          { stack: [
              { text: 'PRECIO MÁXIMO', font: 'PPMori', fontSize: 6.5, characterSpacing: 1.1, color: PALETTE.inkSoft, margin: [0, 0, 0, 7] },
              { text: [{ text: 'US$ ', font: 'PPMori', fontSize: 9, color: PALETTE.inkSoft }, { text: formatMoney(summary.precioMax), font: 'Albra', fontSize: 16, color: PALETTE.copper }], margin: [0, 0, 0, 3] },
              { text: 'Con bodega', font: 'PPMori', fontSize: 7, color: PALETTE.inkSoft }
          ] }
        ]]
      },
      layout: {
        fillColor: function () { return PALETTE.panel; },
        hLineWidth: function () { return 0; },
        vLineWidth: function (i) { return (i === 0 || i === 4) ? 0 : 0.5; },
        vLineColor: function () { return PALETTE.line; },
        paddingLeft: function (i) { return i === 0 ? 18 : 16; },
        paddingRight: function () { return 14; },
        paddingTop: function () { return 16; },
        paddingBottom: function () { return 16; }
      },
      margin: [0, 0, 0, 26]
    };
  }

  function modelTitleCell(modelo, count) {
    return {
      columns: [
        { width: 'auto', text: serifRuns(modelo, 19, PALETTE.copper), margin: [0, 0, 10, 0] },
        { width: '*', text: serifRuns('Modelo ' + modelo, 13, PALETTE.ink), margin: [0, 5, 0, 0] },
        { width: 'auto', text: [
            { text: formatInt(count) + ' ', font: 'PPMori', bold: true, fontSize: 10, color: PALETTE.copper },
            { text: count === 1 ? 'UNIDAD' : 'UNIDADES', font: 'PPMori', fontSize: 7.5, characterSpacing: 1, color: PALETTE.inkSoft }
          ], alignment: 'right', margin: [0, 6, 0, 0] }
      ]
    };
  }

  function columnHeaderRow() {
    var h = function (t, align) {
      return { text: t, font: 'PPMori', fontSize: 6.5, characterSpacing: 1.1, color: PALETTE.inkSoft, alignment: align || 'left' };
    };
    return [h('UNIDAD'), h('NIVEL'), h('MTS² TOTALES'), h('PRECIO CON BODEGA', 'right')];
  }

  function dataRow(rec) {
    return [
      { text: unitLabel(rec), font: 'PPMori', bold: true, fontSize: 9.5, color: PALETTE.ink },
      { text: levelLabel(rec), font: 'PPMori', fontSize: 9, color: PALETTE.inkMid },
      { text: formatArea(rec.mts2) + ' m²', font: 'PPMori', fontSize: 9, color: PALETTE.inkMid },
      { text: [
          { text: 'US$ ', font: 'PPMori', fontSize: 7.5, color: PALETTE.inkSoft },
          { text: formatMoney(rec.precio), font: 'PPMori', bold: true, fontSize: 9.5, color: PALETTE.ink }
        ], alignment: 'right' }
    ];
  }

  function modelBlock(group) {
    var widths = ['*', 70, 96, 124];
    var body = [];
    // Fila 0: título del modelo (colSpan 4)
    body.push([
      { colSpan: 4, margin: [0, 2, 0, 8], stack: [modelTitleCell(group.modelo, group.units.length)] },
      {}, {}, {}
    ]);
    // Fila 1: encabezados de columna
    body.push(columnHeaderRow());
    // Filas de datos
    group.units.forEach(function (rec) { body.push(dataRow(rec)); });

    return {
      table: { headerRows: 2, dontBreakRows: true, widths: widths, body: body },
      layout: {
        fillColor: function () { return null; },
        hLineWidth: function (i) {
          if (i === 0) return 0;          // sin borde superior
          if (i === 1) return 0;          // bajo el título (lo da la fila de headers)
          if (i === 2) return 0.7;        // bajo encabezados de columna
          return 0.5;                     // separadores entre filas de datos
        },
        hLineColor: function (i) { return i === 2 ? PALETTE.line : PALETTE.lineSoft; },
        hLineStyle: function () { return { dash: { length: 1.4, space: 2.4 } }; },
        vLineWidth: function () { return 0; },
        paddingLeft: function (i) { return i === 0 ? 0 : 8; },
        paddingRight: function (i, node) { return i === node.table.widths.length - 1 ? 0 : 8; },
        paddingTop: function (i) { return i <= 1 ? 4 : 7; },
        paddingBottom: function (i) { return i <= 1 ? 6 : 7; }
      },
      margin: [0, 0, 0, 22]
    };
  }

  function sectionHeader() {
    return {
      columns: [
        { width: 248, text: 'Inventario disponible por modelo', font: 'Albra', fontSize: 15, color: PALETTE.ink },
        { width: '*', margin: [12, 11, 0, 0], canvas: [
            { type: 'line', x1: 0, y1: 0, x2: 230, y2: 0, lineWidth: 0.7, lineColor: PALETTE.line, dash: { length: 1.4, space: 2.4 } }
          ] }
      ],
      margin: [0, 0, 0, 4]
    };
  }

  // ---- Header (hero página 1 + banda running pág. 2+) --------------
  function heroHeader(opts) {
    return {
      columns: [
        { width: 150, image: 'aria-logo.png', width: 132, margin: [0, 6, 0, 0] },
        { width: '*', alignment: 'right', stack: [
            { text: 'REPORTE COMERCIAL · DISPONIBILIDAD', font: 'PPMori', fontSize: 6.5, characterSpacing: 1.3, color: PALETTE.inkSoft, margin: [0, 0, 0, 4] },
            { text: 'Apartamentos Disponibles', font: 'Albra', fontSize: 23, color: PALETTE.ink, margin: [0, 0, 0, 5], lineHeight: 1 },
            { text: [
                { text: 'Actualizado ', font: 'PPMori', fontSize: 7.5, color: PALETTE.inkSoft },
                { text: opts.fecha, font: 'PPMori', bold: true, fontSize: 7.5, color: PALETTE.inkMid },
                { text: ' · Precio con bodega · USD', font: 'PPMori', fontSize: 7.5, color: PALETTE.inkSoft }
              ] }
          ] }
      ]
    };
  }

  // ---- docDefinition completo --------------------------------------
  function buildDocDefinition(records, opts) {
    opts = opts || {};
    var fecha = opts.fecha || todayEs();
    var summary = buildSummary(records);
    var groups = groupByModel(records);

    var content = [];
    content.push(heroHeader({ fecha: fecha }));
    content.push({ canvas: [{ type: 'line', x1: 0, y1: 0, x2: 524, y2: 0, lineWidth: 0.8, lineColor: PALETTE.line }], margin: [0, 16, 0, 18] });
    content.push(statBar(summary));
    content.push(sectionHeader());
    content.push({ text: 'Agrupado por modelo · precio con bodega y área total en m²', font: 'PPMori', fontSize: 7.5, color: PALETTE.inkSoft, margin: [0, 0, 0, 18] });
    groups.forEach(function (g) { content.push(modelBlock(g)); });

    return {
      pageSize: 'LETTER',
      pageMargins: [44, 46, 44, 52],
      defaultStyle: { font: 'PPMori', color: PALETTE.ink },
      info: {
        title: 'Aria Escalón · Apartamentos Disponibles',
        author: 'Aria Escalón',
        subject: 'Reporte comercial de disponibilidad'
      },
      background: function (currentPage, pageSize) {
        return { canvas: [{ type: 'rect', x: 0, y: 0, w: pageSize.width, h: pageSize.height, color: PALETTE.paper }] };
      },
      header: function (currentPage) {
        if (currentPage === 1) return null;
        return {
          margin: [44, 22, 44, 0],
          columns: [
            { width: '*', text: [
                { text: 'ARIA ESCALÓN', font: 'PPMori', bold: true, fontSize: 7, characterSpacing: 1, color: PALETTE.copper },
                { text: '  ·  Apartamentos Disponibles', font: 'PPMori', fontSize: 7, color: PALETTE.inkSoft }
              ] },
            { width: 'auto', text: fecha, font: 'PPMori', fontSize: 7, color: PALETTE.inkSoft, alignment: 'right' }
          ]
        };
      },
      footer: function (currentPage, pageCount) {
        return {
          margin: [44, 8, 44, 0],
          columns: [
            { width: '*', text: 'ARIA ESCALÓN · REPORTE COMERCIAL INTERNO', font: 'PPMori', fontSize: 6.5, characterSpacing: 1, color: PALETTE.inkSoft },
            { width: 'auto', text: [
                { text: 'Disponibilidad de inventario', font: 'Albra', fontSize: 8.5, color: PALETTE.inkMid },
                { text: pageCount > 1 ? '   ·   ' + currentPage + ' / ' + pageCount : '', font: 'PPMori', fontSize: 6.5, color: PALETTE.inkSoft }
              ], alignment: 'right' }
          ]
        };
      },
      content: content
    };
  }

  var AriaReport = {
    PALETTE: PALETTE,
    formatMoney: formatMoney,
    formatArea: formatArea,
    formatInt: formatInt,
    todayEs: todayEs,
    normalizeRows: normalizeRows,
    groupByModel: groupByModel,
    buildSummary: buildSummary,
    buildDocDefinition: buildDocDefinition
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = AriaReport;
  root.AriaReport = AriaReport;

})(typeof window !== 'undefined' ? window : this);
