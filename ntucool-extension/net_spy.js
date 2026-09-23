/* NTU COOL 成績與附件小幫手 — 網路請求監看（在「主世界」執行）
 *
 * content script 預設跑在跟頁面隔離的 JS 環境（isolated world），沒辦法直接攔截
 * 頁面自己（NTU COOL / Canvas 的程式碼）發出的 fetch / XMLHttpRequest。
 * 這支檔案透過 manifest.json 裡的 "world": "MAIN" 設定，直接注入到頁面本身的
 * JS 環境執行，攔截跟「送出評論/附件」有關的請求，把「真正的 HTTP 狀態碼與回應內容」
 * 用 CustomEvent 廣播出去，讓 content.js 可以準確判斷送出到底成功還是失敗，
 * 而不是用「留言區數量有沒有變多」這種容易誤判的間接方式猜測。
 *
 * 只讀取跟記錄網路請求的結果，不會修改、攔截或阻擋任何請求本身。
 */
(function () {
  'use strict';
  if (window.__ntucoolNetSpyInstalled) return;
  window.__ntucoolNetSpyInstalled = true;

  const EVENT_NAME = '__ntucoolHelperNetEvent';
  const WATCH_PATTERN = /\/submissions(\/|\.|\?|$)/; // 評分/評論/附件相關的請求都會經過 .../submissions/...

  function emit(detail) {
    try {
      window.dispatchEvent(new CustomEvent(EVENT_NAME, { detail }));
    } catch (e) { /* 忽略 */ }
  }

  // ---- fetch ----
  if (window.fetch) {
    const origFetch = window.fetch;
    window.fetch = function (input, init) {
      const url = typeof input === 'string' ? input : (input && input.url) || '';
      const method = (init && init.method) || (input && input.method) || 'GET';
      const promise = origFetch.apply(this, arguments);
      if (WATCH_PATTERN.test(url)) {
        promise.then((res) => {
          try {
            res.clone().text().then((body) => {
              emit({ url, method, status: res.status, ok: res.ok, body: (body || '').slice(0, 2000), at: Date.now() });
            }).catch(() => {
              emit({ url, method, status: res.status, ok: res.ok, body: '', at: Date.now() });
            });
          } catch (e) { /* 忽略 */ }
        }).catch((err) => {
          emit({ url, method, status: 0, ok: false, body: String(err && err.message || err), at: Date.now() });
        });
      }
      return promise;
    };
  }

  // ---- XMLHttpRequest ----
  const OrigXHR = window.XMLHttpRequest;
  if (OrigXHR) {
    const origOpen = OrigXHR.prototype.open;
    const origSend = OrigXHR.prototype.send;
    OrigXHR.prototype.open = function (method, url) {
      this.__ntucoolMethod = method;
      this.__ntucoolUrl = url;
      return origOpen.apply(this, arguments);
    };
    OrigXHR.prototype.send = function () {
      if (WATCH_PATTERN.test(this.__ntucoolUrl || '')) {
        this.addEventListener('loadend', () => {
          let body = '';
          try { body = (this.responseText || '').slice(0, 2000); } catch (e) { /* 非文字回應，略過 */ }
          emit({ url: this.__ntucoolUrl, method: this.__ntucoolMethod, status: this.status, ok: this.status >= 200 && this.status < 300, body, at: Date.now() });
        });
      }
      return origSend.apply(this, arguments);
    };
  }
})();
