// Lógica pura (sin pantalla ni red) para poder probarla por separado.
(function (raiz) {
  'use strict';

  // Columnas del Excel, en el orden esperado. Solo se llenan las que tienen "campo".
  const COLUMNAS = [
    { clave: 'sep', titulo: 'No. SEP', campo: 'sep' },
    { clave: 'inventario', titulo: 'No. Inventario' },
    { clave: 'area', titulo: 'Area de Trabajo' },
    { clave: 'resguardo', titulo: 'Resguardo', campo: 'resguardo' },
    { clave: 'descripcion', titulo: 'Descripción del articulo', campo: 'descripcion' },
    { clave: 'marca', titulo: 'Marca' },
    { clave: 'modelo', titulo: 'Modelo' },
    { clave: 'serie', titulo: 'Serie', campo: 'serie' },
    { clave: 'observaciones', titulo: 'Observaciones', campo: 'observaciones' },
    { clave: 'estado', titulo: 'Estado', campo: 'estado' },
  ];

  // "Descripción del Artículo " -> "descripcion del articulo"; "No. SEP" -> "no sep"
  function normalizarEncabezado(texto) {
    return String(texto == null ? '' : texto)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/[.:#°º]/g, ' ').replace(/\s+/g, ' ').trim();
  }

  // Ubica cada columna por su encabezado, sin importar orden, acentos o mayúsculas.
  function mapearColumnas(encabezados) {
    const normalizados = encabezados.map(normalizarEncabezado);
    const indices = {};
    const faltantes = [];
    for (const col of COLUMNAS) {
      const i = normalizados.indexOf(normalizarEncabezado(col.titulo));
      if (i === -1) faltantes.push(col.titulo); else indices[col.clave] = i;
    }
    return { indices, faltantes };
  }

  // Forma comparable de un identificador: sin espacios extra, en mayúsculas;
  // si es solo números, sin ceros a la izquierda (Excel a veces los quita).
  function normalizarId(valor) {
    const s = String(valor == null ? '' : valor).replace(/^'/, '').replace(/\s+/g, ' ').trim().toUpperCase();
    if (/^\d+$/.test(s)) return s.replace(/^0+(?=\d)/, '');
    return s;
  }

  // Número de la primera fila de una dirección como "Hoja1!A1:J20" o "'Mi hoja'!$A$5:$J$9".
  function primeraFila(direccion) {
    const m = /!\$?[A-Z]+\$?(\d+)/i.exec(String(direccion || ''));
    return m ? parseInt(m[1], 10) : 1;
  }

  // valores: matriz de la tabla completa (fila 0 = encabezados).
  // Regresa las filas de Excel donde ya existen el No. SEP y/o la Serie.
  function buscarDuplicados(valores, indices, datos, direccion) {
    const filaInicio = primeraFila(direccion);
    const sep = normalizarId(datos.sep);
    const serie = normalizarId(datos.serie);
    const res = { sep: [], serie: [] };
    for (let i = 1; i < valores.length; i++) {
      const fila = valores[i];
      const filaExcel = filaInicio + i;
      if (sep && normalizarId(fila[indices.sep]) === sep) res.sep.push(filaExcel);
      if (serie && normalizarId(fila[indices.serie]) === serie) res.serie.push(filaExcel);
    }
    return res;
  }

  // El apóstrofo hace que Excel guarde el valor como texto: conserva ceros iniciales,
  // no recorta números largos y evita que algo que empiece con "=" se vuelva fórmula.
  function comoTexto(valor) {
    const s = String(valor == null ? '' : valor).trim();
    return s ? "'" + s : '';
  }

  // Arma la fila con la cantidad exacta de columnas de la tabla; las que no llenamos van vacías.
  function construirFila(indices, totalColumnas, datos) {
    const fila = new Array(totalColumnas).fill('');
    for (const col of COLUMNAS) {
      if (col.campo && indices[col.clave] !== undefined) fila[indices[col.clave]] = comoTexto(datos[col.campo]);
    }
    return fila;
  }

  // Revisa que lo que Excel guardó sea exactamente lo enviado (ceros incluidos).
  function verificarFilaGuardada(filaGuardada, indices, datos) {
    const problemas = [];
    for (const col of COLUMNAS) {
      if (!col.campo || indices[col.clave] === undefined) continue;
      const enviado = String(datos[col.campo] == null ? '' : datos[col.campo]).trim();
      const guardado = String(filaGuardada[indices[col.clave]] == null ? '' : filaGuardada[indices[col.clave]]).trim();
      if (enviado !== guardado) problemas.push(`${col.titulo}: se envió "${enviado}" y Excel guardó "${guardado}"`);
    }
    return problemas;
  }

  // Valida el formulario antes de subir.
  function validar(datos, longitudSEP) {
    const errores = [];
    const avisos = [];
    if (!datos.sep && !datos.serie) errores.push('Escanea el código de barras (No. SEP) o escribe la Serie: el registro necesita al menos uno de los dos.');
    if (datos.sep && !/^\d+$/.test(datos.sep)) errores.push('El No. SEP solo puede tener números.');
    if (datos.sep && longitudSEP && datos.sep.length !== longitudSEP) avisos.push(`El No. SEP tiene ${datos.sep.length} dígitos; se esperaban ${longitudSEP}.`);
    if (!datos.resguardo) avisos.push('No se capturó el Resguardo (ubicación).');
    if (!datos.estado) errores.push('Falta el Estado.');
    return { errores, avisos };
  }

  // Del texto leído en la etiqueta toma la línea con letras que está debajo del número.
  // Solo se usa para la descripción; el No. SEP sale únicamente del código de barras.
  function extraerDescripcion(texto) {
    const lineas = String(texto || '').split('\n').map(l => l.trim());
    let ancla = -1;
    lineas.forEach((l, i) => {
      if (ancla !== -1) return;
      const s = l.replace(/[\s.,'`\-_]/g, '');
      if (s.length < 8) return;
      const digitos = (s.match(/\d/g) || []).length;
      const pareceNumero = (s.match(/[\dOoQDIilL|!ZzSs$GBg]/g) || []).length;
      if (digitos >= 3 && pareceNumero / s.length >= 0.85) ancla = i;
    });
    if (ancla === -1) return '';
    for (let i = ancla + 1; i < lineas.length; i++) {
      const l = lineas[i].replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 .,\/-]/g, ' ').replace(/\s+/g, ' ').trim();
      const letras = (l.match(/[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/g) || []).length;
      if (letras >= 4 && letras / l.replace(/\s/g, '').length >= 0.7) return l.toUpperCase();
    }
    return '';
  }

  // Convierte un enlace para compartir en el identificador que pide Graph (/shares/{id}).
  function codificarEnlace(url) {
    const bytes = new TextEncoder().encode(String(url).trim());
    let binario = '';
    bytes.forEach(b => { binario += String.fromCharCode(b); });
    return 'u!' + btoa(binario).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  }

  const Logica = {
    COLUMNAS, normalizarEncabezado, mapearColumnas, normalizarId, primeraFila, buscarDuplicados,
    comoTexto, construirFila, verificarFilaGuardada, validar, extraerDescripcion, codificarEnlace,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = Logica;
  else raiz.Logica = Logica;
})(typeof window !== 'undefined' ? window : globalThis);
