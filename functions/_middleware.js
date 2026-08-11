// Cloudflare Pages Functions：對整個網站做一層簡單的帳號密碼保護 (HTTP Basic Auth)。
// 帳號密碼不寫在程式碼裡，要在 Cloudflare Pages 專案設定裡設環境變數：
//   Settings -> Environment variables -> 新增
//     SITE_USER     = 你要的帳號
//     SITE_PASSWORD = 你要的密碼
// 設定完要重新部署一次（Deployments -> 最新那筆 -> Retry deployment）才會生效。
//
// 本機直接開 index.html 不會經過這層（file:// 本來就只有你電腦看得到），
// 這層只有在部署到 Cloudflare Pages、透過網址瀏覽時才會擋。

function unauthorized() {
  return new Response('請輸入帳號密碼', {
    status: 401,
    headers: { 'WWW-Authenticate': 'Basic realm="restricted"' },
  });
}

export async function onRequest(context) {
  const { request, env, next } = context;
  const user = env.SITE_USER;
  const pass = env.SITE_PASSWORD;

  // 還沒設定帳密環境變數的話，先擋下來並提示，不要就這樣把整個網站放行
  if (!user || !pass) {
    return new Response(
      '網站還沒設定帳號密碼。到 Cloudflare Pages 專案 -> Settings -> Environment variables，' +
      '設定 SITE_USER 和 SITE_PASSWORD 兩個變數，然後重新部署一次。',
      { status: 503 }
    );
  }

  const auth = request.headers.get('Authorization');
  if (!auth || !auth.startsWith('Basic ')) return unauthorized();

  let decoded;
  try {
    decoded = atob(auth.slice(6));
  } catch (e) {
    return unauthorized();
  }
  const idx = decoded.indexOf(':');
  if (idx < 0) return unauthorized();
  const u = decoded.slice(0, idx);
  const p = decoded.slice(idx + 1);
  if (u !== user || p !== pass) return unauthorized();

  return next();
}
