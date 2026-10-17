// 화면 ↔ Apps Script 통신. text/plain 으로 보내면 CORS 사전요청 없이 동작합니다.
(function () {
  class ApiError extends Error {}

  async function api(action, payload) {
    const url = (window.APP_CONFIG || {}).API_URL || '';
    if (!/^https:\/\//.test(url)) {
      throw new ApiError('아직 서버 주소가 설정되지 않았어요. config.js 의 API_URL 을 확인해 주세요.');
    }
    const body = JSON.stringify(Object.assign({ action }, payload || {}));
    let lastErr;
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body,
          redirect: 'follow',
        });
        const json = await res.json();
        if (!json.ok) throw new ApiError(json.error || '요청을 처리하지 못했어요.');
        return json.data;
      } catch (err) {
        if (err instanceof ApiError) throw err;
        lastErr = err;
        await new Promise(r => setTimeout(r, 700 * (attempt + 1)));
      }
    }
    console.error(lastErr);
    throw new ApiError('인터넷 연결이 불안정해요. 잠시 후 다시 시도해 주세요.');
  }

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function toast(msg) {
    let t = document.getElementById('toast');
    if (!t) { t = document.createElement('div'); t.id = 'toast'; t.className = 'toast'; document.body.appendChild(t); }
    t.textContent = msg; t.classList.add('show');
    clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2600);
  }

  const store = {
    get(area, k) { try { return window[area].getItem(k); } catch (e) { return null; } },
    set(area, k, v) { try { window[area].setItem(k, v); } catch (e) {} },
    del(area, k) { try { window[area].removeItem(k); } catch (e) {} },
  };

  const normRoom = v => String(v || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

  window.App = { api, ApiError, esc, toast, store, normRoom };
})();
