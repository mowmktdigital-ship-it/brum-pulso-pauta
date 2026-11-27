/**
 * Pulso de Pauta BRUM — exportación diaria de Google Analytics 4 a este Google Sheet.
 * Gratis: usa Google Apps Script + el servicio avanzado "Google Analytics Data API".
 *
 * INSTALACIÓN (una sola vez, la hace alguien con acceso de lectura a la propiedad GA4 de Brum):
 *  1. Crear un Google Sheet nuevo llamado "GA4 Brum - Pulso".
 *  2. Extensiones → Apps Script. Borrar lo que haya y pegar todo este archivo.
 *  3. En PROPERTY_ID (abajo) poner el ID numérico de la propiedad GA4 de Brum
 *     (GA4 → Administrar → Configuración de la propiedad → "ID de la propiedad").
 *  4. En el editor de Apps Script: "Servicios" (+) → "Google Analytics Data API" → Agregar.
 *  5. Elegir la función "instalarDisparadorDiario" y tocar "Ejecutar". Aceptar los permisos.
 *     Eso crea la actualización automática todos los días a las 13 h y corre una primera vez.
 *  6. Compartir el Sheet con mowmktdigital@gmail.com como Lector y pasarle el link a Agos.
 */

const PROPERTY_ID = 'PONER_ACA_EL_ID_NUMERICO_DE_LA_PROPIEDAD_GA4';
const ZONA = 'America/Argentina/Buenos_Aires';

// Evento de GA4 que cuenta como lead (formulario enviado). Cambiar si en Brum se llama distinto.
const EVENTO_LEAD = 'generate_lead';

function actualizarGA4() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoy = new Date();
  const desde = Utilities.formatDate(new Date(hoy.getFullYear(), hoy.getMonth(), 1), ZONA, 'yyyy-MM-dd');

  // 1) Sesiones por fuente / medio (de dónde llega la gente a la web, incluidas las redes)
  escribirReporte_(ss, 'ga4_fuentes',
    ['sessionSourceMedium'],
    ['sessions', 'totalUsers', 'engagedSessions', 'keyEvents'],
    desde, null);

  // 2) Leads por pieza de contenido (UTM campaign + UTM content), solo el evento de lead
  escribirReporte_(ss, 'ga4_leads_contenido',
    ['sessionSourceMedium', 'sessionCampaignName', 'sessionManualAdContent'],
    ['eventCount'],
    desde, EVENTO_LEAD);

  // 3) Sesiones por pieza de contenido (mismas dimensiones, para calcular tasa)
  escribirReporte_(ss, 'ga4_sesiones_contenido',
    ['sessionSourceMedium', 'sessionCampaignName', 'sessionManualAdContent'],
    ['sessions'],
    desde, null);
}

function escribirReporte_(ss, nombreHoja, dims, mets, desde, filtroEvento) {
  const req = AnalyticsData.newRunReportRequest();
  req.dimensions = dims.map(function (d) { const x = AnalyticsData.newDimension(); x.name = d; return x; });
  req.metrics = mets.map(function (m) { const x = AnalyticsData.newMetric(); x.name = m; return x; });
  const rango = AnalyticsData.newDateRange();
  rango.startDate = desde;
  rango.endDate = 'today';
  req.dateRanges = [rango];
  req.limit = 1000;
  if (filtroEvento) {
    const f = AnalyticsData.newFilter();
    f.fieldName = 'eventName';
    const sf = AnalyticsData.newStringFilter();
    sf.value = filtroEvento;
    sf.matchType = 'EXACT';
    f.stringFilter = sf;
    const fe = AnalyticsData.newFilterExpression();
    fe.filter = f;
    req.dimensionFilter = fe;
  }
  const resp = AnalyticsData.Properties.runReport(req, 'properties/' + PROPERTY_ID);

  const filas = [
    ['actualizado', Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd HH:mm'), 'desde', desde],
    dims.concat(mets)
  ];
  (resp.rows || []).forEach(function (row) {
    filas.push(row.dimensionValues.map(function (v) { return v.value; })
      .concat(row.metricValues.map(function (v) { return Number(v.value); })));
  });
  const ancho = Math.max.apply(null, filas.map(function (f) { return f.length; }));
  filas.forEach(function (f) { while (f.length < ancho) f.push(''); });

  const hoja = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
  hoja.clearContents();
  hoja.getRange(1, 1, filas.length, ancho).setValues(filas);
}

function instalarDisparadorDiario() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'actualizarGA4') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('actualizarGA4').timeBased().everyDays(1).atHour(13).inTimezone(ZONA).create();
  actualizarGA4();
}
