const ALLOWED_METHODS = new Set([
  'obtenerEstadoBloqueoLogin',
  'iniciarSesion',
  'cerrarSesion',
  'listarAuditoria',
  'repararEstructuraHojas',
  'migrarCarrosAValores',
  'listarUsuarios',
  'crearUsuario',
  'eliminarUsuarioPermanentemente',
  'agregarClasificaciones',
  'listarPartesAdmin',
  'actualizarParte',
  'eliminarParte',
  'registrarMovimientoUnificado',
  'listarRegistros',
  'listarHistorial',
  'obtenerResumenHXH',
  'actualizarRegistro',
  'obtenerResumenVisual',
  'obtenerResumenPorDia',
]);

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Método no permitido.' });
  }

  const endpoint = process.env.GAS_WEB_APP_URL;
  const secret = process.env.GAS_SHARED_SECRET;
  if (!endpoint || !secret) {
    return res.status(500).json({ ok: false, error: 'Falta configurar la conexión segura con Apps Script en Vercel.' });
  }

  let requestBody = req.body;
  if (Buffer.isBuffer(requestBody)) {
    requestBody = requestBody.toString('utf8');
  }
  if (typeof requestBody === 'string') {
    try {
      requestBody = JSON.parse(requestBody);
    } catch (error) {
      return res.status(400).json({ ok: false, error: 'El cuerpo de la solicitud no es JSON válido.' });
    }
  }

  if (
    !requestBody ||
    typeof requestBody.method !== 'string' ||
    !ALLOWED_METHODS.has(requestBody.method) ||
    !Array.isArray(requestBody.args) ||
    requestBody.args.length > 8
  ) {
    return res.status(400).json({ ok: false, error: 'La operación solicitada no está habilitada.' });
  }

  try {
    const upstream = await fetch(endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ method: requestBody.method, args: requestBody.args, secret }),
      redirect: 'follow',
    });
    let result;
    try {
      result = await upstream.json();
    } catch (error) {
      return res.status(502).json({
        ok: false,
        error: 'Apps Script no devolvió JSON. Revisa que la aplicación web esté desplegada con acceso para la conexión de Vercel.',
      });
    }

    if (!upstream.ok || !result || typeof result.ok !== 'boolean') {
      return res.status(502).json({ ok: false, error: 'Apps Script devolvió una respuesta no válida.' });
    }
    return res.status(result.ok ? 200 : 400).json(result);
  } catch (error) {
    console.error('No se pudo conectar con Apps Script:', error);
    return res.status(502).json({ ok: false, error: 'No se pudo conectar con Apps Script.' });
  }
};
