// Configuración del programa de inventario. Edita solo los valores entre comillas.
// Instrucciones completas en README.md.
window.CONFIG = {
  // "Id. de aplicación (cliente)" del registro de la app en Microsoft Entra (portal de Azure).
  clientId: 'PEGA-AQUI-EL-ID-DE-APLICACION',

  // "Id. de directorio (inquilino)" de tu organización. Si lo dejas en 'organizations',
  // podrá entrar cualquier cuenta de trabajo o escuela de Microsoft.
  tenantId: 'organizations',

  // Enlace para compartir del Excel en OneDrive o SharePoint (botón Compartir > Copiar vínculo).
  enlaceExcel: 'PEGA-AQUI-EL-ENLACE-DEL-EXCEL',

  // Nombre de la tabla de Excel donde se agregan los registros (Diseño de tabla > Nombre de la tabla).
  // Si lo dejas vacío se usa la primera tabla del libro.
  tabla: '',

  // Dígitos que debe tener el No. SEP (código de barras). Solo se usa para avisar.
  longitudSEP: 16,
};
