// Cortex Claude proxy — Cloudflare Worker
// The browser posts { model, max_tokens, system, messages } to this Worker.
// The Worker attaches the server-side API key and forwards to Anthropic.
// Deploy on Cloudflare Workers (free tier) and set ANTHROPIC_KEY as a secret.
//
//   wrangler secret put ANTHROPIC_KEY
//   wrangler deploy
//
// Or via dashboard: Workers & Pages → Create → paste this file → Settings →
// Variables → Add variable (encrypt) → ANTHROPIC_KEY = sk-ant-...

// Comma-separated list of origins allowed to call this worker.
// Lock this down to your deployed site(s) so random people can't drain your key.
const ALLOWED_ORIGINS = [
  'https://steventyoo.github.io',
  'http://localhost:8000',
  'http://127.0.0.1:8000',
  'null', // file:// origins send "null"
];

function corsHeaders(origin) {
  const allow = ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0];
  return {
    'access-control-allow-origin': allow,
    'access-control-allow-methods': 'POST, OPTIONS',
    'access-control-allow-headers': 'content-type',
    'access-control-max-age': '86400',
    'vary': 'origin',
  };
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get('origin') || '';
    const cors = corsHeaders(origin);

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: cors });
    }

    if (request.method !== 'POST') {
      return new Response('cortex proxy: POST only', { status: 405, headers: cors });
    }

    if (!env.ANTHROPIC_KEY) {
      return new Response(
        JSON.stringify({ error: { message: 'Worker missing ANTHROPIC_KEY secret' } }),
        { status: 500, headers: { ...cors, 'content-type': 'application/json' } }
      );
    }

    let body;
    try {
      body = await request.text();
      JSON.parse(body); // sanity-check it's JSON
    } catch {
      return new Response(
        JSON.stringify({ error: { message: 'Invalid JSON body' } }),
        { status: 400, headers: { ...cors, 'content-type': 'application/json' } }
      );
    }

    const upstream = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': env.ANTHROPIC_KEY,
        'anthropic-version': '2023-06-01',
      },
      body,
    });

    return new Response(upstream.body, {
      status: upstream.status,
      headers: { ...cors, 'content-type': 'application/json' },
    });
  },
};
