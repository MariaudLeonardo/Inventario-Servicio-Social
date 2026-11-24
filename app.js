// Inventario de muebles: escaneo de código de barras + registro en Excel compartido (Microsoft Graph).
// La app SOLO agrega filas: no existe ninguna función que edite o borre registros del Excel.
(function () {
  'use strict';
  const C = window.CONFIG || {};
  const L = window.Logica;
  const $ = id => document.getElementById(id);
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const SCOPES = ['Files.ReadWrite.All'];
  const URL_APP = new URL('./', location.href).href; // https://usuario.github.io/repositorio/
  const CLAVE_HISTORIAL = 'inventario_subidos_v2';
  const CLAVE_BORRADOR = 'inventario_borrador_v2';
  const CAMPOS = ['sep', 'resguardo', 'descripcion', 'serie', 'observaciones', 'estado'];

  let msalApp = null;
  let libro = null;        // { driveId, itemId, nombre, tablaId, tablaNombre }
  let ocrWorker = null;    // Tesseract, solo para la descripción
  let detector = null;     // BarcodeDetector
  let lectura = 0;         // evita que una lectura vieja sobrescriba una nueva
  let subiendo = false;

  // ---------- Utilidades de pantalla ----------
  function mensaje(id, tipo, html) { $(id).innerHTML = html ? `<div class="${tipo}">${html}</div>` : ''; }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function leerFormulario() {
    const d = {};
    CAMPOS.forEach(c => { d[c] = $(c).value.trim(); });
    return d;
  }
  function guardarBorrador() {
    try { localStorage.setItem(CLAVE_BORRADOR, JSON.stringify(leerFormulario())); } catch (e) { /* sin espacio */ }
  }
  function restaurarBorrador() {
    try {
      const d = JSON.parse(localStorage.getItem(CLAVE_BORRADOR) || 'null');
      if (d) CAMPOS.forEach(c => { if (d[c] !== undefined && d[c] !== '') $(c).value = d[c]; });
    } catch (e) { /* borrador dañado: se ignora */ }
  }
  function limpiarFormulario() {
    ['sep', 'descripcion', 'serie', 'observaciones'].forEach(c => { $(c).value = ''; });
    $('estado').value = 'Bueno';
    delete $('descripcion').dataset.auto;
    $('preview').style.display = 'none';
    mensaje('estadoFoto', '', '');
    mensaje('avisoSep', '', '');
    guardarBorrador(); // conserva el Resguardo para el siguiente mueble
  }
  function historial() {
    try { return JSON.parse(localStorage.getItem(CLAVE_HISTORIAL) || '[]'); } catch (e) { return []; }
  }
  function pintarHistorial() {
    const h = historial();
    $('contador').textContent = `${h.length} subido${h.length === 1 ? '' : 's'} desde este celular`;
    $('tabla').innerHTML = h.slice(0, 100).map(r =>
      `<tr><td class="num">${esc(r.sep) || '—'}</td><td>${esc(r.serie) || '—'}</td><td>${esc(r.resguardo)}</td><td>${esc(r.descripcion)}</td><td>${esc(r.fila)}</td></tr>`
    ).join('');
    const opciones = campo => [...new Set(h.map(r => r[campo]).filter(Boolean))].map(v => `<option value="${esc(v)}"></option>`).join('');
    $('listaResguardo').innerHTML = opciones('resguardo');
    $('listaDesc').innerHTML = opciones('descripcion');
  }

  // ---------- Inicio de sesión (MSAL) ----------
  function configuracionCompleta() {
    return C.clientId && !/PEGA-AQUI/.test(C.clientId) && C.enlaceExcel && !/PEGA-AQUI/.test(C.enlaceExcel);
  }

  async function iniciarMsal() {
    if (!window.msal) throw new Error('No se pudo cargar la librería de inicio de sesión de Microsoft. Revisa la conexión a internet.');
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: C.clientId,
        authority: 'https://login.microsoftonline.com/' + (C.tenantId || 'organizations'),
        redirectUri: URL_APP,
        postLogoutRedirectUri: URL_APP,
      },
      cache: { cacheLocation: 'localStorage' },
    });
    await msalApp.initialize();
    const resp = await msalApp.handleRedirectPromise();
    if (resp && resp.account) msalApp.setActiveAccount(resp.account);
    else if (!msalApp.getActiveAccount() && msalApp.getAllAccounts().length) msalApp.setActiveAccount(msalApp.getAllAccounts()[0]);
  }

  function pintarSesion() {
    const cuenta = msalApp && msalApp.getActiveAccount();
    $('usuario').textContent = cuenta ? (cuenta.name || cuenta.username) : 'Sin sesión';
    $('btnEntrar').style.display = cuenta ? 'none' : 'inline-block';
    $('btnSalir').style.display = cuenta ? 'inline-block' : 'none';
    $('btnSubir').disabled = !cuenta;
    $('btnProbar').disabled = !cuenta;
  }

  // Regresa un token, o null si hubo que redirigir a la página de Microsoft (el borrador queda guardado).
  async function obtenerToken() {
    const cuenta = msalApp.getActiveAccount();
    if (!cuenta) { guardarBorrador(); await msalApp.loginRedirect({ scopes: SCOPES }); return null; }
    try {
      return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account: cuenta })).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        guardarBorrador();
        await msalApp.acquireTokenRedirect({ scopes: SCOPES, account: cuenta });
        return null;
      }
      throw e;
    }
  }

  // ---------- Microsoft Graph ----------
  function errorGraph(estado, detalle) {
    const textos = {
      400: 'Solicitud rechazada por Excel.',
      401: 'La sesión expiró. Vuelve a iniciar sesión.',
      403: 'Tu cuenta no tiene permiso de edición sobre el Excel compartido.',
      404: 'No se encontró el Excel o la tabla. Revisa el enlace y el nombre de la tabla en config.js.',
      409: 'Excel está ocupado con otro cambio. Intenta de nuevo.',
      423: 'El Excel está bloqueado. Intenta de nuevo en un momento.',
      429: 'Demasiadas solicitudes a Microsoft. Espera unos segundos e intenta de nuevo.',
    };
    const base = textos[estado] || (estado >= 500 ? 'El servicio de Microsoft no respondió. Intenta de nuevo.' : `Error ${estado}.`);
    const err = new Error(base + (detalle ? ` (${detalle})` : ''));
    err.estado = estado;
    return err;
  }

  async function graph(token, metodo, ruta, cuerpo) {
    // Solo las consultas (GET) se reintentan; un POST nunca se repite solo para no duplicar filas.
    const intentos = metodo === 'GET' ? 3 : 1;
    for (let n = 1; ; n++) {
      let r;
      try {
        r = await fetch(GRAPH + ruta, {
          method: metodo,
          headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: cuerpo ? JSON.stringify(cuerpo) : undefined,
        });
      } catch (e) {
        if (n < intentos) { await esperar(1000 * n); continue; }
        throw new Error('Sin conexión a internet o Microsoft no respondió.');
      }
      if (r.ok) return r.status === 204 ? null : r.json();
      if (n < intentos && [429, 500, 502, 503, 504].includes(r.status)) {
        const espera = parseInt(r.headers.get('Retry-After') || '0', 10) * 1000 || 1000 * n;
        await esperar(Math.min(espera, 10000));
        continue;
      }
      let detalle = '';
      try { const j = await r.json(); detalle = (j.error && (j.error.message || j.error.code)) || ''; } catch (e) { /* sin cuerpo */ }
      throw errorGraph(r.status, detalle);
    }
  }
  const esperar = ms => new Promise(ok => setTimeout(ok, ms));

  async function abrirLibro(token) {
    if (libro) return libro;
    const item = await graph(token, 'GET', `/shares/${L.codificarEnlace(C.enlaceExcel)}/driveItem?$select=id,name,parentReference`);
    if (!/\.xlsx$/i.test(item.name || '')) throw new Error(`El enlace apunta a "${item.name}", que no es un archivo .xlsx.`);
    const base = `/drives/${encodeURIComponent(item.parentReference.driveId)}/items/${encodeURIComponent(item.id)}/workbook`;
    const tablas = (await graph(token, 'GET', `${base}/tables?$select=id,name`)).value || [];
    if (!tablas.length) throw new Error('El Excel no tiene ninguna tabla. Selecciona los encabezados y usa Insertar > Tabla (ver README).');
    const tabla = C.tabla
      ? tablas.find(t => t.name.toLowerCase() === String(C.tabla).trim().toLowerCase())
      : tablas[0];
    if (!tabla) throw new Error(`No existe la tabla "${C.tabla}". Tablas en el libro: ${tablas.map(t => t.name).join(', ')}.`);
    libro = { base, nombre: item.name, tablaRuta: `${base}/tables/${encodeURIComponent(tabla.id)}`, tablaNombre: tabla.name };
    return libro;
  }

  async function leerTabla(token) {
    const lb = await abrirLibro(token);
    const rango = await graph(token, 'GET', `${lb.tablaRuta}/range?$select=address,values`);
    const valores = rango.values || [];
    const { indices, faltantes } = L.mapearColumnas(valores[0] || []);
    if (faltantes.length) throw new Error(`A la tabla "${lb.tablaNombre}" le faltan columnas: ${faltantes.join(', ')}.`);
    return { lb, valores, indices, direccion: rango.address };
  }

  // ---------- Subir registro ----------
  async function subir() {
    if (subiendo) return;
    const datos = leerFormulario();
    const { errores, avisos } = L.validar(datos, C.longitudSEP);
    if (errores.length) { mensaje('estadoSubida', 'aviso', errores.map(esc).join('<br>')); return; }
    if (avisos.length && !confirm(avisos.join('\n') + '\n\n¿Subir de todos modos?')) return;

    subiendo = true;
    $('btnSubir').disabled = true;
    mensaje('estadoSubida', 'info', 'Revisando si ya existe en el Excel...');
    try {
      const token = await obtenerToken();
      if (!token) return; // redirigiendo a Microsoft
      const { lb, valores, indices, direccion } = await leerTabla(token);

      const dup = L.buscarDuplicados(valores, indices, datos, direccion);
      const repetidos = [];
      if (dup.sep.length) repetidos.push(`El No. SEP ${esc(datos.sep)} ya existe en el Excel (fila ${dup.sep.join(', ')}).`);
      if (dup.serie.length) repetidos.push(`La Serie ${esc(datos.serie)} ya existe en el Excel (fila ${dup.serie.join(', ')}).`);
      if (repetidos.length) {
        mensaje('estadoSubida', 'aviso', repetidos.join('<br>') + '<br>No se agregó nada.');
        return;
      }

      mensaje('estadoSubida', 'info', 'Agregando al Excel...');
      const fila = L.construirFila(indices, valores[0].length, datos);
      const resp = await graph(token, 'POST', `${lb.tablaRuta}/rows/add`, { index: null, values: [fila] });

      const guardada = (resp && resp.values && resp.values[0]) || [];
      const problemas = L.verificarFilaGuardada(guardada, indices, datos);
      const filaExcel = (resp && typeof resp.index === 'number') ? L.primeraFila(direccion) + 1 + resp.index : '';

      const h = historial();
      h.unshift({ ...datos, fila: filaExcel, fecha: new Date().toLocaleString('es-MX') });
      localStorage.setItem(CLAVE_HISTORIAL, JSON.stringify(h.slice(0, 500)));
      limpiarFormulario();
      pintarHistorial();

      if (problemas.length) {
        mensaje('estadoSubida', 'aviso', `Se agregó en la fila ${filaExcel}, pero revisa en el Excel: ${problemas.map(esc).join('; ')}`);
      } else {
        mensaje('estadoSubida', 'ok', `✓ Agregado al Excel "${esc(lb.nombre)}"${filaExcel ? ` en la fila ${filaExcel}` : ''}.`);
      }
    } catch (e) {
      if (e.estado === 401) libro = null;
      mensaje('estadoSubida', 'aviso', esc(e.message) +
        '<br>Los datos siguen en el formulario. Si vuelves a intentar y dice que ya existe, es que sí se guardó.');
    } finally {
      subiendo = false;
      pintarSesion();
    }
  }

  async function probarConexion() {
    mensaje('estadoSubida', 'info', 'Probando conexión con el Excel...');
    try {
      const token = await obtenerToken();
      if (!token) return;
      libro = null;
      const { lb, valores } = await leerTabla(token);
      const registros = valores.slice(1).filter(f => f.some(v => String(v).trim() !== '')).length;
      mensaje('estadoSubida', 'ok', `✓ Conectado a "${esc(lb.nombre)}", tabla "${esc(lb.tablaNombre)}". Columnas correctas. Registros actuales: ${registros}.`);
    } catch (e) {
      mensaje('estadoSubida', 'aviso', esc(e.message));
    }
  }

  // ---------- Código de barras + descripción ----------
  function cargarImagen(url) {
    return new Promise((ok, mal) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => mal(new Error('No se pudo abrir la foto.')); i.src = url; });
  }

  // Escala a ~2000px, escala de grises y más contraste.
  function prepararImagen(img) {
    const escala = Math.min(1, 2000 / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement('canvas');
    c.width = Math.round(img.naturalWidth * escala); c.height = Math.round(img.naturalHeight * escala);
    const ctx = c.getContext('2d');
    ctx.drawImage(img, 0, 0, c.width, c.height);
    const d = ctx.getImageData(0, 0, c.width, c.height);
    for (let i = 0; i < d.data.length; i += 4) {
      let g = 0.299 * d.data[i] + 0.587 * d.data[i + 1] + 0.114 * d.data[i + 2];
      g = Math.max(0, Math.min(255, (g - 128) * 1.6 + 128));
      d.data[i] = d.data[i + 1] = d.data[i + 2] = g;
    }
    ctx.putImageData(d, 0, 0);
    return c;
  }

  async function leerCodigoBarras(img) {
    if (!('BarcodeDetector' in window)) throw new Error('Este navegador no pudo cargar el lector de códigos de barras.');
    if (!detector) detector = new BarcodeDetector();
    for (const fuente of [img, prepararImagen(img)]) {
      const valores = (await detector.detect(fuente))
        .map(r => String(r.rawValue || '').replace(/\D/g, ''))
        .filter(v => v.length >= 8);
      const num = valores.find(v => v.length === C.longitudSEP) || valores[0];
      if (num) return num;
    }
    return '';
  }

  // Solo se ejecuta después de leer el código de barras, y solo llena la descripción.
  async function llenarDescripcion(img, miLectura) {
    const campo = $('descripcion');
    if (campo.value.trim() && !campo.dataset.auto) return; // el usuario ya escribió algo
    if (!window.Tesseract) return;
    mensaje('estadoDesc', 'info', 'Leyendo la descripción...');
    try {
      if (!ocrWorker) ocrWorker = Tesseract.createWorker('spa');
      const w = await ocrWorker;
      const { data } = await w.recognize(prepararImagen(img));
      if (miLectura !== lectura) return; // ya se tomó otra foto
      const desc = L.extraerDescripcion(data.text);
      if (desc && (!campo.value.trim() || campo.dataset.auto)) {
        campo.value = desc;
        campo.dataset.auto = '1';
        guardarBorrador();
        mensaje('estadoDesc', 'ok', '✓ Descripción llenada automáticamente. Corrígela si hace falta.');
      } else {
        mensaje('estadoDesc', '', '');
      }
    } catch (e) {
      ocrWorker = null;
      mensaje('estadoDesc', '', '');
    }
  }

  async function alTomarFoto(e) {
    const archivo = e.target.files[0];
    e.target.value = '';
    if (!archivo) return;
    const miLectura = ++lectura;
    const url = URL.createObjectURL(archivo);
    $('preview').src = url;
    $('preview').style.display = 'block';
    mensaje('estadoDesc', '', '');
    mensaje('estadoFoto', 'info', 'Leyendo código de barras...');
    try {
      const img = await cargarImagen(url);
      const numero = await leerCodigoBarras(img);
      if (miLectura !== lectura) return;
      if (!numero) {
        mensaje('estadoFoto', 'aviso', 'No se pudo leer el código de barras. Toma otra foto (barras completas, derechas y sin brillo) o escribe el No. SEP a mano. Si el mueble no tiene código, déjalo vacío y llena la Serie.');
        return;
      }
      $('sep').value = numero;
      validarSep();
      guardarBorrador();
      mensaje('estadoFoto', 'ok', '✓ No. SEP leído del código de barras.');
      llenarDescripcion(img, miLectura);
    } catch (err) {
      mensaje('estadoFoto', 'aviso', esc(err.message));
    }
  }

  function validarSep() {
    const v = $('sep').value;
    mensaje('avisoSep', 'aviso', v && C.longitudSEP && v.length !== C.longitudSEP ? `Tiene ${v.length} dígitos; se esperaban ${C.longitudSEP}.` : '');
  }

  // ---------- Arranque ----------
  async function iniciar() {
    restaurarBorrador();
    pintarHistorial();
    $('camara').addEventListener('change', alTomarFoto);
    $('sep').addEventListener('input', () => { $('sep').value = $('sep').value.replace(/\D/g, ''); validarSep(); });
    $('descripcion').addEventListener('input', () => { delete $('descripcion').dataset.auto; mensaje('estadoDesc', '', ''); });
    $('formulario').addEventListener('input', guardarBorrador);
    $('formulario').addEventListener('change', guardarBorrador);
    $('btnLimpiar').addEventListener('click', () => { if (confirm('¿Limpiar el formulario?')) limpiarFormulario(); });
    $('btnSubir').addEventListener('click', subir);
    $('btnProbar').addEventListener('click', probarConexion);
    $('btnEntrar').addEventListener('click', () => {
      if (!msalApp) { mensaje('estadoSubida', 'aviso', 'El inicio de sesión no está disponible: revisa config.js y la conexión a internet.'); return; }
      guardarBorrador();
      msalApp.loginRedirect({ scopes: SCOPES }).catch(e => mensaje('estadoSubida', 'aviso', esc(e.message)));
    });
    $('btnSalir').addEventListener('click', () => {
      if (msalApp && confirm('¿Cerrar sesión?')) msalApp.logoutRedirect({ account: msalApp.getActiveAccount(), postLogoutRedirectUri: URL_APP });
    });
    validarSep();

    if (!configuracionCompleta()) {
      mensaje('estadoSubida', 'aviso', 'Falta configurar config.js (clientId y enlaceExcel). Ver README.md.');
      pintarSesion();
      return;
    }
    try {
      await iniciarMsal();
    } catch (e) {
      mensaje('estadoSubida', 'aviso', 'Error al iniciar sesión: ' + esc(e.message));
    }
    pintarSesion();
  }

  iniciar();
})();
