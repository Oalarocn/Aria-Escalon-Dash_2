/* =====================================================================
   Aria Escalón · Lógica de la app web (sin backend)
   ---------------------------------------------------------------------
   - Registra fuentes/logo (base64) en pdfMake.
   - Lee el Excel en el navegador con SheetJS.
   - Normaliza, genera la vista previa y permite descargar el PDF.
   Todo ocurre en el cliente; no se envía nada a ningún servidor.
   ===================================================================== */
(function () {
  'use strict';

  // ---- Registro de fuentes y logo en pdfMake ------------------------
  function registerPdfAssets() {
    if (!window.pdfMake) { console.error('pdfmake no cargó'); return false; }
    pdfMake.vfs = window.ARIA_VFS || {};
    pdfMake.fonts = {
      Albra: {
        normal: 'Albra-Light.otf', bold: 'Albra-Light.otf',
        italics: 'Albra-Light.otf', bolditalics: 'Albra-Light.otf'
      },
      PPMori: {
        normal: 'PPMori-Regular.otf', bold: 'PPMori-SemiBold.otf',
        italics: 'PPMori-Regular.otf', bolditalics: 'PPMori-SemiBold.otf'
      }
    };
    return true;
  }

  // ---- Datos de ejemplo (para el botón "Ver ejemplo") --------------
  var SAMPLE_ROWS = [
    ['Modelo', 'Unidad', 'Nivel', 'Mts2 Totales', 'Precio con Bodega'],
    ['A', 1, 2, 136.38, 406603.90], ['D', 2, 2, 77.91, 252533.58], ['D', 3, 2, 77.91, 252533.58],
    ['B', 5, 2, 117.90, 319725.00], ['E', 6, 2, 111.35, 304500.00], ['E', 7, 2, 104.79, 304709.67],
    ['A', 1, 3, 135.01, 406452.29], ['D', 2, 3, 87.99, 267658.91], ['A', 4, 3, 135.21, 406511.23],
    ['B', 9, 3, 116.90, 322604.29], ['D', 2, 4, 82.32, 268634.50], ['B', 5, 4, 108.32, 320617.06],
    ['B', 9, 4, 108.37, 320766.08], ['A', 1, 5, 135.01, 414382.65], ['A', 1, 6, 136.38, 422465.41],
    ['D', 3, 6, 83.08, 275712.45], ['B', 5, 6, 107.49, 327211.81], ['B', 5, 7, 108.86, 334422.86],
    ['B', 9, 7, 116.90, 335036.39], ['A', 4, 7, 135.21, 427247.39], ['C+', 6, 8, 93.94, 261498.89],
    ['C', 7, 8, 73.24, 237695.45], ['C+', 6, 10, 93.94, 266381.41]
  ];

  // ---- Estado -------------------------------------------------------
  var state = { records: [], currentDoc: null, sourceName: '' };

  // ---- Utilidades de fecha -----------------------------------------
  function isoToday() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function fechaEsFromInput(value) {
    if (!value) return window.AriaReport.todayEs();
    var parts = value.split('-');
    if (parts.length !== 3) return window.AriaReport.todayEs();
    var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    return window.AriaReport.todayEs(d);
  }
  function fileStamp(value) {
    return value || isoToday();
  }

  // ---- Diagnóstico de librerías ------------------------------------
  // Si alguna no cargó (típico en GitHub Pages sin .nojekyll), lo decimos
  // con nombre y archivo, en vez de fallar en silencio.
  function missingLibs() {
    var miss = [];
    if (!window.XLSX)      miss.push('js/lib/xlsx.full.min.js (lectura de Excel)');
    if (!window.pdfMake)   miss.push('js/lib/pdfmake.min.js (generación de PDF)');
    if (!window.ARIA_VFS)  miss.push('js/lib/aria-assets.js (fuentes y logo)');
    if (!window.AriaReport) miss.push('js/report.js (diseño del reporte)');
    return miss;
  }

  // ---- DOM refs (se asignan en init) -------------------------------
  var el = {};

  function setStatus(msg, type) {
    if (!el.status) return;
    el.status.textContent = msg || '';
    el.status.className = 'status' + (type ? ' status--' + type : '');
    el.status.style.display = msg ? 'block' : 'none';
  }

  // ---- Procesar registros normalizados -----------------------------
  function applyRecords(records, sourceName) {
    state.records = records;
    state.sourceName = sourceName || '';
    var s = window.AriaReport.buildSummary(records);
    renderSummary(s);
    generatePreview();
    el.workspace.classList.add('is-active');
    el.downloadBtn.disabled = false;
    setStatus(records.length + ' apartamentos cargados correctamente.', 'ok');
  }

  function renderSummary(s) {
    var chips = [
      { k: 'Apartamentos', v: window.AriaReport.formatInt(s.total) },
      { k: 'Modelos', v: s.modelosList.join(' · ') },
      { k: 'Niveles', v: window.AriaReport.formatInt(s.niveles) },
      { k: 'Precio mín.', v: 'US$ ' + window.AriaReport.formatMoney(s.precioMin) },
      { k: 'Precio máx.', v: 'US$ ' + window.AriaReport.formatMoney(s.precioMax) }
    ];
    el.summary.innerHTML = chips.map(function (c) {
      return '<div class="chip"><span class="chip__k">' + c.k + '</span>' +
             '<span class="chip__v">' + c.v + '</span></div>';
    }).join('');
  }

  // ---- Generar documento / vista previa / descarga -----------------
  function currentFecha() {
    return fechaEsFromInput(el.dateInput && el.dateInput.value);
  }
  function isSinglePage() {
    return !!(el.singlePage && el.singlePage.checked);
  }

  // Construye el docDefinition y lo entrega por callback.
  // En modo "una sola página", primero mide la altura real del contenido
  // (render de medición) y luego arma la página a esa altura exacta.
  function withDoc(cb) {
    var fecha = currentFecha();
    if (!isSinglePage()) {
      cb(window.AriaReport.buildDocDefinition(state.records, { fecha: fecha }));
      return;
    }
    try {
      var measureDd = window.AriaReport.buildMeasureDoc(state.records, { fecha: fecha });
      pdfMake.createPdf(measureDd).getBuffer(function () {
        var H = measureDd.computePageHeight();
        if (!H) {
          // El contenido excede el alto máximo de una sola página: aviso y multipágina.
          setStatus('Hay demasiados datos para una sola página; se generó en varias hojas.', 'info');
        }
        cb(window.AriaReport.buildDocDefinition(state.records, { fecha: fecha, pageHeight: H || undefined }));
      });
    } catch (e) {
      console.error(e);
      cb(window.AriaReport.buildDocDefinition(state.records, { fecha: fecha }));
    }
  }

  function generatePreview() {
    el.preview.classList.add('is-loading');
    withDoc(function (dd) {
      try {
        state.currentDoc = pdfMake.createPdf(dd);
        state.currentDoc.getDataUrl(function (url) {
          el.previewFrame.src = url;
          el.preview.classList.remove('is-loading');
        });
      } catch (e) {
        console.error(e);
        setStatus('No se pudo generar la vista previa: ' + e.message, 'error');
        el.preview.classList.remove('is-loading');
      }
    });
  }

  function downloadPdf() {
    if (!state.records.length) return;
    withDoc(function (dd) {
      try {
        var name = 'Aria_Escalon_Disponibilidad_' + fileStamp(el.dateInput && el.dateInput.value) +
                   (isSinglePage() ? '_continuo' : '') + '.pdf';
        pdfMake.createPdf(dd).download(name);
      } catch (e) {
        console.error(e);
        setStatus('No se pudo descargar el PDF: ' + e.message, 'error');
      }
    });
  }

  // ---- Lectura del Excel con SheetJS -------------------------------
  function readWorkbook(arrayBuffer, sourceName) {
    try {
      var wb = XLSX.read(arrayBuffer, { type: 'array' });
      var firstSheet = wb.SheetNames[0];
      var sheet = wb.Sheets[firstSheet];
      var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
      var result = window.AriaReport.normalizeRows(rows);
      if (result.errors && result.errors.length) {
        setStatus(result.errors[0], 'error');
        return;
      }
      if (!result.records.length) {
        setStatus('No se encontraron filas de datos en el archivo.', 'error');
        return;
      }
      applyRecords(result.records, sourceName);
    } catch (e) {
      console.error(e);
      setStatus('No se pudo leer el archivo. ¿Es un Excel válido (.xlsx / .csv)? ' + e.message, 'error');
    }
  }

  function handleFile(file) {
    if (!file) return;
    var miss = missingLibs();
    if (miss.length) {
      setStatus('La app cargó, pero no se pudieron cargar estas librerías: ' + miss.join(' · ') +
                '. En GitHub Pages esto suele resolverse creando un archivo vacío llamado .nojekyll en la raíz del repositorio.', 'error');
      return;
    }
    var okExt = /\.(xlsx|xlsm|xls|csv)$/i.test(file.name);
    if (!okExt) { setStatus('Formato no soportado. Usa .xlsx, .xls o .csv', 'error'); return; }
    setStatus('Leyendo "' + file.name + '"…', 'info');
    var reader = new FileReader();
    reader.onload = function (e) { readWorkbook(e.target.result, file.name); };
    reader.onerror = function () { setStatus('Error al leer el archivo.', 'error'); };
    reader.readAsArrayBuffer(file);
  }

  // ---- Init ---------------------------------------------------------
  function init() {
    el.dropzone = document.getElementById('dropzone');
    el.fileInput = document.getElementById('fileInput');
    el.sampleBtn = document.getElementById('sampleBtn');
    el.status = document.getElementById('status');
    el.summary = document.getElementById('summary');
    el.workspace = document.getElementById('workspace');
    el.preview = document.getElementById('preview');
    el.previewFrame = document.getElementById('previewFrame');
    el.downloadBtn = document.getElementById('downloadBtn');
    el.dateInput = document.getElementById('dateInput');
    el.singlePage = document.getElementById('singlePage');

    if (el.dateInput) el.dateInput.value = isoToday();

    // --- 1) Cablear SIEMPRE la interfaz (clic, arrastre, ejemplo) ----
    //     Aunque falte una librería, el uploader queda activo y damos
    //     un mensaje claro al intentar usarlo.

    // Click en dropzone abre el selector
    el.dropzone.addEventListener('click', function () { el.fileInput.click(); });
    el.dropzone.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.fileInput.click(); }
    });
    el.fileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
    });

    // Drag & drop
    ['dragenter', 'dragover'].forEach(function (ev) {
      el.dropzone.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        el.dropzone.classList.add('is-drag');
      });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      el.dropzone.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        el.dropzone.classList.remove('is-drag');
      });
    });
    el.dropzone.addEventListener('drop', function (e) {
      var dt = e.dataTransfer;
      if (dt && dt.files && dt.files[0]) handleFile(dt.files[0]);
    });

    // Botón ejemplo
    el.sampleBtn.addEventListener('click', function () {
      var miss = missingLibs();
      if (miss.length) {
        setStatus('No se pudieron cargar: ' + miss.join(' · ') +
                  '. En GitHub Pages, crea un archivo vacío llamado .nojekyll en la raíz del repositorio y vuelve a recargar.', 'error');
        return;
      }
      setStatus('Cargando datos de ejemplo…', 'info');
      var result = window.AriaReport.normalizeRows(SAMPLE_ROWS);
      applyRecords(result.records, 'ejemplo');
    });

    // Regenerar al cambiar fecha
    if (el.dateInput) {
      el.dateInput.addEventListener('change', function () {
        if (state.records.length) generatePreview();
      });
    }

    // Regenerar al alternar "una sola página"
    if (el.singlePage) {
      el.singlePage.addEventListener('change', function () {
        if (state.records.length) generatePreview();
      });
    }

    el.downloadBtn.addEventListener('click', downloadPdf);

    // --- 2) Verificar librerías y registrar recursos del PDF ---------
    var miss = missingLibs();
    if (miss.length) {
      setStatus('La interfaz está activa, pero faltan recursos: ' + miss.join(' · ') +
                '. Si esto ocurre solo en GitHub Pages, crea un archivo vacío llamado .nojekyll en la raíz del repositorio y recarga con Ctrl/Cmd+Shift+R.', 'error');
      return;
    }
    registerPdfAssets();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})();
