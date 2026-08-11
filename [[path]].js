
const GAS_URL = "https://script.google.com/macros/s/AKfycbxH3Q99G0dM0xxWwKSOO3Lds9HJnPi0UG_JCpzHXRAEx3o2tf18mI0Zfws7T0bjsKlNLg/exec";
const PAGE_MAP = {"/": "Dashboard", "/dashboard": "Dashboard", "/requests": "Requests", "/admin": "Admin", "/statistics": "Statistics", "/assistant": "Assistant"};

export async function onRequest(context) {
  const request = context.request;
  const url = new URL(request.url);
  const path = url.pathname.replace(/\/+$/, '') || '/';

  // RPC endpoint: browser -> Cloudflare -> Apps Script.
  if (path === '/api/rpc') {
    if (request.method !== 'POST') {
      return new Response('Method Not Allowed', { status: 405 });
    }
    const body = await request.text();

    const upstream = await fetch(GAS_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body
    });

    return new Response(await upstream.text(), {
      status: upstream.status,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store'
      }
    });
  }

  // Voice recorder is hosted directly on Pages so microphone permission
  // belongs to the user's domain instead of script.google.com.
  if (path === '/voice' || path === '/voice.html') {
    const assetUrl = new URL('/voice.html', url);
    return context.env.ASSETS.fetch(new Request(assetUrl, request));
  }

  // Resolve the Apps Script page.
  let page = PAGE_MAP[path];
  if (!page && path === '/') {
    page = url.searchParams.get('page') || 'Dashboard';
  }
  if (!page) {
    return new Response('Not Found', { status: 404 });
  }

  const upstreamUrl = new URL(GAS_URL);
  upstreamUrl.searchParams.set('page', page);

  // Preserve authentication/session token and supported query parameters.
  for (const [key, value] of url.searchParams.entries()) {
    if (key === 'page') continue;
    upstreamUrl.searchParams.append(key, value);
  }

  const upstream = await fetch(upstreamUrl.toString(), {
    method: 'GET',
    redirect: 'follow',
    headers: {
      'Accept': 'text/html,application/xhtml+xml'
    }
  });

  let body = await upstream.text();

  // Rewrite ALL Apps Script deployment URLs to this Pages origin.
  // This prevents sidebar/navigation links from sending the browser to script.google.com.
  body = body.split(GAS_URL).join(url.origin);

  // Normal absolute Apps Script URLs.
  body = body.replace(
    /https:\/\/script\.google\.com\/macros\/s\/[^"'\s<>]+\/exec/gi,
    url.origin
  );

  // Escaped URLs that may appear inside JavaScript strings.
  body = body.replace(
    /https:\\/\\/script\\.google\\.com\\/macros\\/s\\/[^"'\s<>]+\\/exec/gi,
    url.origin
  );

  // Force the app's internal reload base to remain on Cloudflare Pages.
  body = body.replace(
    /window\.__WEBAPP_BASE_URL__\s*=\s*(['"])[^'"]*\1/,
    'window.__WEBAPP_BASE_URL__ = ' + JSON.stringify(url.origin)
  );

  // If no recorder URL was configured yet, use the Pages-hosted voice page.
  body = body.replace(
    /window\.__VOICE_RECORDER_URL__\s*=\s*"";/,
    'window.__VOICE_RECORDER_URL__ = ' + JSON.stringify(url.origin + '/voice') + ';'
  );

  // Remove the Apps Script warning banner if Google included it in the HTML.
  body = body.replace(
    /<table[^>]*>[\s\S]*?This application was created by a Google Apps Script user[\s\S]*?<\/table>/gi,
    ''
  );

  // The original HTML uses google.script.run. Replace it with our same-origin RPC bridge.
  body = body.split('google.script.run').join('window.__GAS_RUN__');

  const injection = "\n<script>\n(function () {\n  // Local replacement for google.script.run.\n  // The browser talks only to the current Cloudflare domain.\n  function makeRunner(success, failure, userObject) {\n    const target = {};\n    return new Proxy(target, {\n      get: function (_target, prop) {\n        if (prop === 'withSuccessHandler') {\n          return function (cb) { return makeRunner(cb, failure, userObject); };\n        }\n        if (prop === 'withFailureHandler') {\n          return function (cb) { return makeRunner(success, cb, userObject); };\n        }\n        if (prop === 'withUserObject') {\n          return function (obj) { return makeRunner(success, failure, obj); };\n        }\n        if (prop === 'then') return undefined;\n        return function () {\n          const args = Array.prototype.slice.call(arguments);\n          fetch('/api/rpc', {\n            method: 'POST',\n            headers: {'Content-Type': 'application/json'},\n            body: JSON.stringify({ action: String(prop), args: args })\n          })\n          .then(function (r) {\n            return r.text().then(function (text) {\n              let data;\n              try { data = JSON.parse(text); }\n              catch (e) { throw new Error('\u0627\u0633\u062a\u062c\u0627\u0628\u0629 \u063a\u064a\u0631 \u0635\u0627\u0644\u062d\u0629 \u0645\u0646 \u0627\u0644\u062e\u0627\u062f\u0645.'); }\n              if (!r.ok || !data.ok) {\n                const err = new Error((data.error && data.error.message) || '\u062d\u062f\u062b \u062e\u0637\u0623 \u0641\u064a \u0627\u0644\u062e\u0627\u062f\u0645.');\n                err.name = (data.error && data.error.name) || 'Error';\n                throw err;\n              }\n              return data.result;\n            });\n          })\n          .then(function (result) {\n            if (typeof success === 'function') success(result, userObject);\n          })\n          .catch(function (err) {\n            if (typeof failure === 'function') failure(err, userObject);\n            else console.error(err);\n          });\n        };\n      }\n    });\n  }\n\n  window.__GAS_RUN__ = makeRunner(null, null, null);\n  // Keep a compatibility object for any remaining code that checks google.script.run.\n  window.google = window.google || {};\n  window.google.script = window.google.script || {};\n  window.google.script.run = window.__GAS_RUN__;\n})();\n</script>\n";
  body = body.replace(/<head([^>]*)>/i, '<head$1>' + injection);

  return new Response(body, {
    status: upstream.status,
    headers: {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'no-store, no-cache, must-revalidate'
    }
  });
}
