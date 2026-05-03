// Vercel serverless function — proxy para a Anthropic API.
// A chave nunca sai do servidor. CORS limitado ao domínio do próprio Vercel + localhost.

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';
const ANTHROPIC_VERSION = '2023-06-01';
const DEFAULT_MAX_TOKENS = 4096;
const ALLOWED_MODELS = new Set(['claude-opus-4-5', 'claude-sonnet-4-5']);

function setCors(req, res) {
  const origin = req.headers.origin || '';
  const allowed =
    origin.endsWith('.vercel.app') ||
    origin === 'https://vivword.vercel.app' ||
    origin.startsWith('http://localhost') ||
    origin.startsWith('https://localhost') ||
    // Office hosts iframe the taskpane — these origins call /api/chat.
    origin.endsWith('.officeapps.live.com') ||
    origin.endsWith('.office.com') ||
    origin.endsWith('.office365.com');
  if (allowed) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('Access-Control-Max-Age', '86400');
}

async function readJsonBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  return await new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (chunk) => {
      data += chunk;
      if (data.length > 5_000_000) reject(new Error('Payload demasiado grande'));
    });
    req.on('end', () => {
      if (!data) return resolve({});
      try {
        resolve(JSON.parse(data));
      } catch (e) {
        reject(new Error('JSON inválido'));
      }
    });
    req.on('error', reject);
  });
}

module.exports = async (req, res) => {
  setCors(req, res);

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Método não suportado. Usar POST.' });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: 'ANTHROPIC_API_KEY não configurada no Vercel.' });
    return;
  }

  let body;
  try {
    body = await readJsonBody(req);
  } catch (e) {
    res.status(400).json({ error: e.message });
    return;
  }

  const {
    messages,
    model = 'claude-opus-4-5',
    system,
    max_tokens,
    stream = false,
  } = body || {};

  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: 'Campo "messages" obrigatório (array não vazio).' });
    return;
  }
  if (!ALLOWED_MODELS.has(model)) {
    res.status(400).json({ error: `Modelo não permitido: ${model}` });
    return;
  }

  const payload = {
    model,
    max_tokens: Math.min(Math.max(parseInt(max_tokens, 10) || DEFAULT_MAX_TOKENS, 1), 8192),
    messages,
    stream: !!stream,
  };
  if (system && typeof system === 'string' && system.trim()) {
    payload.system = system;
  }

  let upstream;
  try {
    upstream = await fetch(ANTHROPIC_URL, {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': ANTHROPIC_VERSION,
      },
      body: JSON.stringify(payload),
    });
  } catch (e) {
    res.status(502).json({ error: `Falha ao contactar Anthropic: ${e.message}` });
    return;
  }

  if (!upstream.ok) {
    let detail = '';
    try {
      detail = await upstream.text();
    } catch (_) {}
    let parsed;
    try {
      parsed = JSON.parse(detail);
    } catch (_) {
      parsed = { raw: detail };
    }
    res.status(upstream.status).json({
      error: parsed?.error?.message || `Erro ${upstream.status} da Anthropic`,
      detail: parsed,
    });
    return;
  }

  if (stream) {
    res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Accel-Buffering', 'no');
    res.flushHeaders?.();

    const reader = upstream.body.getReader();
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        if (value) res.write(Buffer.from(value));
      }
      res.end();
    } catch (e) {
      try {
        res.write(`event: error\ndata: ${JSON.stringify({ error: e.message })}\n\n`);
      } catch (_) {}
      res.end();
    }
    return;
  }

  let data;
  try {
    data = await upstream.json();
  } catch (e) {
    res.status(502).json({ error: 'Resposta inválida da Anthropic.' });
    return;
  }
  res.status(200).json(data);
};
