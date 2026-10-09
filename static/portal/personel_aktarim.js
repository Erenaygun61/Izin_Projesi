(() => {
  'use strict';
  const panel = document.getElementById('personnel-transfer');
  if (!panel) return;
  const ui = Object.fromEntries(['status', 'percent', 'progress', 'count', 'time', 'message', 'errors', 'pause', 'resume', 'retry', 'refresh'].map(key => [key, document.getElementById('transfer-' + key)]));
  const base = panel.dataset.jobBase.replace(/baslat\/$/, '');
  let job = null;
  let busy = false;
  let pause = false;
  let began = 0;
  let initialProcessed = 0;
  let pendingStart = null;
  let initializing = !!panel.dataset.lastId;
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const requestId = () => {
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    bytes[6] = (bytes[6] & 15) | 64;
    bytes[8] = (bytes[8] & 63) | 128;
    const hex = Array.from(bytes, value => value.toString(16).padStart(2, '0')).join('');
    return [hex.slice(0, 8), hex.slice(8, 12), hex.slice(12, 16), hex.slice(16, 20), hex.slice(20)].join('-');
  };
  const csrf = () => document.querySelector('input[name="csrfmiddlewaretoken"]').value;
  function lock(value) {
    busy = value;
    document.querySelectorAll('[data-personel-aktarim] button[type="submit"]').forEach(button => { button.disabled = value; });
    ui.pause.hidden = !value;
    ui.pause.disabled = !job || job.done;
    ui.resume.hidden = value || !!job?.done;
    ui.retry.disabled = value;
    ui.refresh.disabled = value;
  }
  function render(data) {
    job = data;
    panel.hidden = false;
    ui.percent.textContent = '%' + data.percent;
    ui.progress.value = data.percent;
    ui.count.textContent = `${data.processed} / ${data.total} kay\u0131t; ${data.success} ba\u015far\u0131l\u0131, ${data.failed} hata`;
    ui.status.textContent = data.done ? (data.failed ? 'Hatalarla Tamamland\u0131' : 'Aktar\u0131m Tamamland\u0131') : 'Personel Aktar\u0131m\u0131';
    const handled = data.processed - initialProcessed;
    const seconds = began && handled > 0 ? Math.ceil((performance.now() - began) / 1000 / handled * (data.total - data.processed)) : null;
    ui.time.textContent = data.done || seconds === null ? '' : `Tahmini kalan: ${seconds < 60 ? seconds + ' sn' : Math.ceil(seconds / 60) + ' dk'}`;
    ui.retry.hidden = !data.done || !data.failed;
    ui.refresh.hidden = !data.done;
    ui.errors.replaceChildren();
    (data.errors || []).forEach(error => {
      const li = document.createElement('li');
      li.textContent = `${error.tip} / ${error.source_id}: ${error.message}`;
      ui.errors.appendChild(li);
    });
    ui.errors.hidden = !data.failed;
  }
  async function api(url, body) {
    const response = await fetch(url, {method: body ? 'POST' : 'GET', body, headers: body ? {'X-CSRFToken': csrf()} : {}, credentials: 'same-origin'});
    const data = await response.json().catch(() => ({}));
    if (!response.ok || response.redirected) {
      const error = new Error(data.error || (response.redirected || response.status === 401 ? 'Oturumunuz sona erdi. Yeniden giri\u015f yap\u0131n.' : 'Ba\u011flant\u0131 kurulamad\u0131. Yeniden deneyin.'));
      error.status = response.status;
      throw error;
    }
    return data;
  }
  async function run() {
    if (busy || !job || job.done) return;
    pause = false;
    began = performance.now();
    initialProcessed = job.processed;
    lock(true);
    ui.message.textContent = 'Aktar\u0131l\u0131yor...';
    try {
      let attempts = 0;
      while (!job.done && !pause) {
        const body = new FormData();
        body.set('processed', job.processed);
        try {
          render(await api(base + job.id + '/adim/', body));
          attempts = 0;
        } catch (error) {
          if (![401, 403, 404].includes(error.status) && attempts < 3) {
            attempts++;
            ui.message.textContent = 'Ba\u011flant\u0131 bekleniyor...';
            await wait(1000 * attempts);
            // A lost response may already have committed: obtain the server cursor first.
            render(await api(base + job.id + '/'));
            continue;
          }
          throw error;
        }
        await wait(30);
      }
      ui.message.textContent = job.done ? (job.failed ? 'Aktar\u0131lamayan kay\u0131tlar a\u015fa\u011f\u0131da listelendi.' : 'T\u00fcm kay\u0131tlar aktar\u0131ld\u0131.') : 'Aktar\u0131m duraklat\u0131ld\u0131.';
      if (job.done) await refreshList();
    } catch (error) {
      ui.message.textContent = error.message;
    } finally { lock(false); }
  }
  async function start(body) {
    if (busy || initializing) return;
    pendingStart = body;
    panel.hidden = false;
    lock(true);
    ui.message.textContent = 'Aktar\u0131m haz\u0131rlan\u0131yor...';
    try {
      render(await api(panel.dataset.startUrl, body));
      pendingStart = null;
      lock(false);
      await run();
    } catch (error) {
      lock(false);
      ui.resume.hidden = false;
      ui.message.textContent = error.message;
    }
  }
  async function refreshList() {
    try {
      const response = await fetch(location.href, {credentials: 'same-origin'});
      if (!response.ok || response.redirected) return;
      const doc = new DOMParser().parseFromString(await response.text(), 'text/html');
      ['.source-stats', '.local-type-tabs', '.local-personnel-list'].forEach(selector => {
        const replacement = doc.querySelector(selector);
        if (replacement) document.querySelector(selector)?.replaceWith(replacement);
      });
    } catch (_) { ui.refresh.hidden = false; }
  }
  document.addEventListener('submit', event => {
    if (!event.target.matches('[data-personel-aktarim]')) return;
    event.preventDefault();
    if (busy || initializing) return;
    const body = new FormData(event.target);
    body.set('islem', event.submitter?.value || 'secilenleri_guncelle');
    body.set('request_id', requestId());
    start(body);
  });
  ui.pause.addEventListener('click', () => { pause = true; ui.message.textContent = 'Mevcut par\u00e7a tamamland\u0131ktan sonra duraklat\u0131lacak.'; });
  ui.resume.addEventListener('click', async () => {
    if (pendingStart) return start(pendingStart);
    try { render(await api(base + job.id + '/')); await run(); } catch (error) { ui.message.textContent = error.message; }
  });
  ui.retry.addEventListener('click', () => {
    const body = new FormData(); body.set('retry', job.id); body.set('request_id', requestId()); start(body);
  });
  ui.refresh.addEventListener('click', () => location.reload());
  if (initializing) {
    lock(true);
    api(base + panel.dataset.lastId + '/').then(data => {
      render(data);
      ui.message.textContent = data.done ? 'Son aktar\u0131m sonucu.' : 'Tamamlanmam\u0131\u015f aktar\u0131m var.';
    }).catch(() => {}).finally(() => { initializing = false; lock(false); });
  }
})();
