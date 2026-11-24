# Inventario de muebles

Página web para el celular que escanea el código de barras de las etiquetas y **agrega** cada mueble a un Excel compartido en OneDrive o SharePoint, usando el inicio de sesión de Microsoft (MSAL) y Microsoft Graph.

- El **No. SEP** sale únicamente del código de barras (o se escribe a mano). Es opcional: hay muebles sin código.
- La **Descripción del artículo** se llena sola, pero solo después de leer el código de barras, y nunca sobrescribe lo que se escribió a mano.
- Antes de agregar, la app revisa en el Excel que **no exista ya** el mismo No. SEP ni la misma **Serie**.
- La app **solo agrega filas**: no tiene ninguna función para editar o borrar registros del Excel.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La pantalla |
| `app.js` | Escaneo, inicio de sesión y envío a Excel |
| `logica.js` | Reglas: columnas, duplicados, validaciones |
| `config.js` | **Lo único que se edita**: IDs de Microsoft y enlace del Excel |

Los archivos `servidor.ps1` e `Iniciar servidor.bat` eran de la versión anterior para red local. Ya no se usan y no se suben a GitHub.

## 1. Preparar el Excel

1. Crea o abre el `.xlsx` en OneDrive o SharePoint.
2. En la fila 1 escribe estos encabezados (el orden puede variar; acentos y mayúsculas no importan):

   `No. SEP` · `No. Inventario` · `Area de Trabajo` · `Resguardo` · `Descripción del articulo` · `Marca` · `Modelo` · `Serie` · `Observaciones` · `Estado`

3. Selecciona los encabezados y usa **Insertar > Tabla** con "La tabla tiene encabezados" marcado.
4. En **Diseño de tabla > Nombre de la tabla**, ponle un nombre, por ejemplo `Inventario`.
5. Recomendado: selecciona las columnas **No. SEP** y **Serie** y dales formato **Texto**, para que Excel nunca quite los ceros iniciales.
6. Comparte el archivo con **permiso de edición** a cada persona que vaya a capturar, y copia el enlace (**Compartir > Copiar vínculo**).
7. **No protejas la hoja**: con la protección activada, Excel rechaza que se agreguen filas desde la app.

## 2. Publicar en GitHub Pages

1. Crea un repositorio en GitHub, por ejemplo `inventario`.
2. Sube `index.html`, `app.js`, `logica.js`, `config.js` y este `README.md`.
3. Ve a **Settings > Pages**, elige **Deploy from a branch**, rama `main`, carpeta `/ (root)` y guarda.
4. Después de un par de minutos la app queda en `https://TU-USUARIO.github.io/inventario/`.

## 3. Registrar la app en Microsoft Entra

1. Entra a <https://entra.microsoft.com> (o al portal de Azure) con una cuenta de tu organización.
2. Ve a **Aplicaciones > Registros de aplicaciones > Nuevo registro**.
   - Nombre: `Inventario de muebles`.
   - Tipos de cuenta: **solo las cuentas de este directorio organizativo**.
   - URI de redirección: plataforma **Aplicación de página única (SPA)**, con la dirección exacta de GitHub Pages **con la diagonal final**, por ejemplo `https://TU-USUARIO.github.io/inventario/`.
3. En **Permisos de API > Agregar un permiso > Microsoft Graph > Permisos delegados**, agrega **Files.ReadWrite.All**.
   - Si tu organización no permite que los usuarios den su propio consentimiento, un administrador debe pulsar **Conceder consentimiento de administrador**.
4. En **Información general**, copia el **Id. de aplicación (cliente)** y el **Id. de directorio (inquilino)**.

## 4. Llenar `config.js`

```js
window.CONFIG = {
  clientId: 'el Id. de aplicación (cliente)',
  tenantId: 'el Id. de directorio (inquilino)',
  enlaceExcel: 'el enlace que copiaste al compartir el Excel',
  tabla: 'Inventario',
  longitudSEP: 16,
};
```

Sube el `config.js` actualizado al repositorio.

## 5. Probar

1. Abre la dirección de GitHub Pages en el celular y toca **Iniciar sesión**.
2. Toca **Probar conexión con el Excel**. Debe aparecer el nombre del archivo, de la tabla y la cantidad de registros.
3. Escanea una etiqueta y confirma que el No. SEP sea idéntico al número impreso, con todo y ceros.

## Notas importantes

- **Quién puede borrar.** La app nunca borra ni edita, pero quien tenga permiso de edición puede modificar el Excel abriéndolo directamente. Para recuperar algo borrado por error, usa el **Historial de versiones** del archivo en OneDrive o SharePoint.
- **Repositorio público.** Si el repositorio es público, cualquiera puede ver `config.js`. El Id. de aplicación no es secreto, y el enlace del Excel solo funciona para cuentas con permiso sobre el archivo.
- **Dos celulares a la vez.** Si dos personas suben el mismo mueble en el mismo segundo, ambas revisiones pueden pasar antes de que exista la fila. Es muy poco probable, pero conviene revisar duplicados al cerrar el inventario.
- **Si un envío falla.** Los datos se quedan en el formulario. Si al reintentar dice "ya existe", el registro sí se guardó la primera vez.
