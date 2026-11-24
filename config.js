// Configuración del programa de inventario. Edita solo los valores entre comillas.
// Instrucciones completas en README.md.
window.CONFIG = {
  // "Id. de aplicación (cliente)" del registro de la app en Microsoft Entra (portal de Azure).
  clientId: '1c7a08b5-d4ac-48a5-9f31-f415eb3077ce',

  // "Id. de directorio (inquilino)" de tu organización. Si lo dejas en 'organizations',
  // podrá entrar cualquier cuenta de trabajo o escuela de Microsoft.
  tenantId: 'organizations',

  // Enlace para compartir del Excel en OneDrive o SharePoint (botón Compartir > Copiar vínculo).
  enlaceExcel: 'https://itleon-my.sharepoint.com/:x:/r/personal/22240327_leon_tecnm_mx/_layouts/15/doc2.aspx?sourcedoc=%7B7FE16355-BFE3-48C8-A1B5-A0CFFA6756E5%7D&file=Book.xlsx&action=editNew&mobileredirect=true&wdPreviousSession=3974d862-855b-4cd7-945b-2baa2298ab9c&wdPreviousSessionSrc=AppHomeWeb',

  // Nombre de la tabla de Excel donde se agregan los registros (Diseño de tabla > Nombre de la tabla).
  // Si lo dejas vacío se usa la primera tabla del libro.
  tabla: '',

  // Dígitos que debe tener el No. SEP (código de barras). Solo se usa para avisar.
  longitudSEP: 16,
};
