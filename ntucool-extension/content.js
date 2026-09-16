/* NTU COOL 成績與附件小幫手 (content script)
 * 只在 SpeedGrader 頁面執行。所有資料都留在瀏覽器本機記憶體/儲存空間，
 * 不會送到任何第三方伺服器；上傳成績/附件仍然是透過你自己登入的 NTU COOL 頁面完成。
 */
(function () {
  'use strict';
  if (window.__ntucoolHelperLoaded) return;
  window.__ntucoolHelperLoaded = true;

  const STORAGE_KEY = 'ntucoolHelper_scoreMapV1';

  const state = {
    scoreRows: [],        // 2D array from Excel/CSV (含表頭)
    idColIdx: -1,
    scoreColIdx: -1,
    scoreMap: {},          // { 'B12345678': '1.8' }
    pdfMap: {},            // { 'B12345678': File }
    pdfUnmatched: [],      // 檔名列表（找不到學號）
    log: [],               // { time, id, name, action, ok, detail }
    autoAdvance: false,
    commentText: '',
  };

  // ---------------- 小工具 ----------------
  function wait(ms) { return new Promise((r) => setTimeout(r, ms)); }

  function extractStudentId(text) {
    if (!text) return null;
    const m = String(text).toUpperCase().match(/\b([A-Z]\d{8})\b/);
    return m ? m[1] : null;
  }

  function nativeSetValue(el, value) {
    const proto = el.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype : window.HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
    setter.call(el, value);
  }

  function fireEvent(el, type, opts) {
    el.dispatchEvent(new Event(type, Object.assign({ bubbles: true, cancelable: true }, opts || {})));
  }

  function fireKey(el, type, key) {
    el.dispatchEvent(new KeyboardEvent(type, { key, code: key, keyCode: 13, which: 13, bubbles: true, cancelable: true }));
  }

  function fmtTime() {
    const d = new Date();
    return d.toLocaleTimeString('zh-TW', { hour12: false });
  }

  // ---------------- SpeedGrader DOM ----------------
  function getStudentSelect() { return document.getElementById('students_selectmenu'); }

  function getCurrentStudent() {
    const sel = getStudentSelect();
    if (!sel || sel.selectedIndex < 0) return null;
    const opt = sel.options[sel.selectedIndex];
    if (!opt) return null;
    const raw = (opt.textContent || '').trim();
    return { raw, id: extractStudentId(raw), optionValue: opt.value };
  }

  function getGradeInput() { return document.getElementById('grading-box-extended'); }

  function setGrade(value) {
    const input = getGradeInput();
    if (!input) return false;
    input.focus();
    nativeSetValue(input, String(value));
    fireEvent(input, 'input');
    fireEvent(input, 'change');
    fireKey(input, 'keydown', 'Enter');
    fireKey(input, 'keyup', 'Enter');
    input.blur();
    return true;
  }

  async function attachFile(file) {
    const addBtn = document.getElementById('add_attachment');
    if (!addBtn) throw new Error('找不到「文檔附件」按鈕，請確認目前在作業評論區塊');
    const before = document.querySelectorAll('input[type=file]').length;
    addBtn.click();
    let input = null;
    for (let i = 0; i < 30; i++) {
      const inputs = document.querySelectorAll('input[type=file]');
      if (inputs.length > before) { input = inputs[inputs.length - 1]; break; }
      await wait(100);
    }
    if (!input) throw new Error('點擊附加檔案後找不到新的上傳欄位（頁面結構可能已變更）');
    const dt = new DataTransfer();
    dt.items.add(file);
    input.files = dt.files;
    fireEvent(input, 'change');
    return true;
  }

  function setCommentText(text) {
    if (!text) return true;
    try {
      if (window.tinymce) {
        const editor = window.tinymce.get('comment_rce_textarea') || (window.tinymce.editors && window.tinymce.editors[0]);
        if (editor) { editor.setContent(text); return true; }
      }
    } catch (e) { /* 忽略，改走下面 fallback */ }
    const ta = document.getElementById('comment_rce_textarea') || document.querySelector('textarea[name=comment]');
    if (ta) {
      nativeSetValue(ta, text);
      fireEvent(ta, 'input');
      fireEvent(ta, 'change');
      return true;
    }
    return false;
  }

  function getSubmitButton() { return document.getElementById('comment_submit_button'); }

  async function submitComment() {
    const btn = getSubmitButton();
    if (!btn) throw new Error('找不到評論區的「提交」按鈕');
    const before = document.querySelectorAll('#comments .comment').length;
    btn.click();
    for (let i = 0; i < 50; i++) {
      const now = document.querySelectorAll('#comments .comment').length;
      if (now > before) return true;
      await wait(200);
    }
    return false; // 逾時：不確定是否成功，需自行確認畫面
  }

  function goNext() {
    const btn = document.getElementById('next-student-button');
    if (btn) { btn.click(); return true; }
    return false;
  }

  // ---------------- Excel 解析 ----------------
  function parseWorkbookFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('讀取檔案失敗'));
      reader.onload = () => {
        try {
          const data = new Uint8Array(reader.result);
          const wb = window.XLSX.read(data, { type: 'array' });
          const sheetName = wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
          resolve({ sheetName, rows });
        } catch (e) { reject(e); }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function guessColumns(headerRow) {
    let idIdx = -1, scoreIdx = -1;
    headerRow.forEach((h, i) => {
      const s = String(h || '');
      if (idIdx === -1 && /(學號|student ?id|studentid|帳號)/i.test(s)) idIdx = i;
      if (scoreIdx === -1 && /(成績|分數|得分|score|grade|points?|total|合計|總分|總計)/i.test(s)) scoreIdx = i;
    });
    // 找不到明確的「成績」欄位名稱時（例如表頭是 Total/P1~P6 這種分項），
    // 退而求其次猜「最右邊有資料的欄位」，並避免跟學號欄位選到同一欄
    if (scoreIdx === -1) {
      for (let i = headerRow.length - 1; i >= 0; i--) {
        if (i !== idIdx && String(headerRow[i] || '').trim() !== '') { scoreIdx = i; break; }
      }
    }
    if (scoreIdx === idIdx) {
      const alt = headerRow.findIndex((h, i) => i !== idIdx && String(h || '').trim() !== '');
      if (alt !== -1) scoreIdx = alt;
    }
    return { idIdx, scoreIdx };
  }

  function buildScoreMap() {
    const map = {};
    const { scoreRows, idColIdx, scoreColIdx } = state;
    if (idColIdx < 0 || scoreColIdx < 0) return map;
    for (let r = 1; r < scoreRows.length; r++) {
      const row = scoreRows[r];
      if (!row) continue;
      const idRaw = row[idColIdx];
      const scoreRaw = row[scoreColIdx];
      const id = extractStudentId(idRaw) || (idRaw ? String(idRaw).trim().toUpperCase() : null);
      if (!id) continue;
      if (scoreRaw === '' || scoreRaw === undefined || scoreRaw === null) continue;
      map[id] = String(scoreRaw).trim();
    }
    return map;
  }

  // ---------------- PDF 資料夾索引 ----------------
  function indexPdfFiles(fileList) {
    const map = {};
    const unmatched = [];
    Array.from(fileList).forEach((f) => {
      if (!/\.pdf$/i.test(f.name)) return;
      const id = extractStudentId(f.name);
      if (id) {
        map[id] = f; // 若同學號有多個檔案，取最後一個（通常代表資料夾內較新的版本）
      } else {
        unmatched.push(f.name);
      }
    });
    return { map, unmatched };
  }

  // ---------------- chrome.storage 持久化（僅存成績表，不存 PDF） ----------------
  function saveScoreMapToStorage() {
    try {
      chrome.storage && chrome.storage.local && chrome.storage.local.set({
        [STORAGE_KEY]: { scoreMap: state.scoreMap, savedAt: Date.now() },
      });
    } catch (e) { /* 忽略儲存失敗，不影響主要功能 */ }
  }

  function loadScoreMapFromStorage() {
    return new Promise((resolve) => {
      try {
        chrome.storage.local.get([STORAGE_KEY], (res) => {
          const data = res && res[STORAGE_KEY];
          resolve(data && data.scoreMap ? data.scoreMap : null);
        });
      } catch (e) { resolve(null); }
    });
  }

  // ---------------- UI ----------------
  let ui = {};

  function buildPanel() {
    const host = document.createElement('div');
    host.id = 'ntucool-helper-host';
    host.style.all = 'initial';
    document.body.appendChild(host);
    const shadow = host.attachShadow({ mode: 'open' });

    const style = document.createElement('style');
    style.textContent = `
      :host { all: initial; }
      .panel {
        position: fixed; top: 90px; right: 16px; width: 320px; max-height: 82vh;
        background: #fff; border: 1px solid #c7cdd6; border-radius: 10px;
        box-shadow: 0 6px 24px rgba(0,0,0,.18); z-index: 2147483647;
        font-family: -apple-system, "Segoe UI", "PingFang TC", "Microsoft JhengHei", sans-serif;
        font-size: 13px; color: #1b2733; display: flex; flex-direction: column; overflow: hidden;
      }
      .panel.collapsed .body { display: none; }
      .head {
        background: #0b3d78; color: #fff; padding: 8px 10px; display: flex; align-items: center;
        justify-content: space-between; cursor: move; user-select: none;
      }
      .head b { font-size: 13px; }
      .head .btns button { background: transparent; border: none; color: #fff; cursor: pointer; font-size: 14px; margin-left: 6px; }
      .body { padding: 10px; overflow-y: auto; }
      section { margin-bottom: 10px; padding-bottom: 10px; border-bottom: 1px dashed #dbe1e8; }
      section:last-child { border-bottom: none; }
      h4 { margin: 0 0 6px; font-size: 12px; color: #0b3d78; }
      .row { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-wrap: wrap; }
      select, input[type=text], textarea {
        width: 100%; box-sizing: border-box; padding: 4px 6px; border: 1px solid #c7cdd6; border-radius: 5px; font-size: 12px;
      }
      textarea { resize: vertical; min-height: 40px; }
      label.small { font-size: 11px; color: #55606b; display:block; margin-bottom:2px; }
      button.act {
        background: #0b3d78; color: #fff; border: none; border-radius: 6px; padding: 6px 10px;
        font-size: 12px; cursor: pointer;
      }
      button.act.secondary { background: #eef2f7; color: #0b3d78; }
      button.act.warn { background: #b3261e; }
      button.act:disabled { opacity: .5; cursor: not-allowed; }
      .stat { background: #f4f7fb; border-radius: 6px; padding: 8px; font-size: 12px; line-height: 1.6; }
      .stat b { color: #0b3d78; }
      .ok { color: #147a3b; }
      .warn { color: #b3261e; }
      .muted { color: #7a8592; }
      .log { max-height: 130px; overflow-y: auto; font-size: 11px; background: #fafbfc; border:1px solid #e5e9ee; border-radius:6px; padding:6px; }
      .log div { padding: 2px 0; border-bottom: 1px solid #eef1f4; }
      .toggle { display:flex; align-items:center; gap:6px; font-size:12px; }
      .grow { flex: 1; }
      .badge { display:inline-block; padding:1px 6px; border-radius:10px; font-size:10px; background:#e5edf7; color:#0b3d78; }
    `;
    shadow.appendChild(style);

    const panel = document.createElement('div');
    panel.className = 'panel';
    panel.innerHTML = `
      <div class="head">
        <b>NTU COOL 成績小幫手</b>
        <span class="btns">
          <button id="ncMin" title="收合">—</button>
        </span>
      </div>
      <div class="body">
        <section>
          <h4>① 成績 Excel / CSV</h4>
          <div class="row"><input type="file" id="ncExcelFile" accept=".xlsx,.xls,.csv" /></div>
          <div class="row" id="ncColRow" style="display:none">
            <div class="grow"><label class="small">學號欄位</label><select id="ncIdCol"></select></div>
            <div class="grow"><label class="small">成績欄位</label><select id="ncScoreCol"></select></div>
          </div>
          <div class="muted" id="ncExcelStatus">尚未載入</div>
        </section>
        <section>
          <h4>② 已批改 PDF 資料夾</h4>
          <div class="row"><input type="file" id="ncPdfFolder" webkitdirectory directory multiple /></div>
          <div class="muted" id="ncPdfStatus">尚未載入</div>
        </section>
        <section>
          <h4>③ 目前學生</h4>
          <div class="stat" id="ncCurrentStat">請開啟 SpeedGrader 並選擇一位學生</div>
        </section>
        <section>
          <h4>④ 評論文字（可留空）</h4>
          <textarea id="ncCommentText" placeholder="例如：已批改，請見附件"></textarea>
        </section>
        <section>
          <div class="row">
            <button class="act secondary grow" id="ncFillGrade">只填成績</button>
            <button class="act secondary grow" id="ncAttachOnly">只附加PDF</button>
          </div>
          <div class="row">
            <button class="act grow" id="ncDoBoth">填成績＋附加並送出</button>
          </div>
          <div class="row toggle">
            <input type="checkbox" id="ncAutoNext" /> <label for="ncAutoNext">成功送出後自動跳下一位</label>
          </div>
          <div class="row">
            <button class="act warn grow" id="ncAutoRun">批次處理剩下所有已比對到的學生</button>
          </div>
        </section>
        <section>
          <h4>紀錄</h4>
          <div class="log" id="ncLog"><div class="muted">尚無紀錄</div></div>
        </section>
      </div>
    `;
    shadow.appendChild(panel);

    // 拖曳
    let dragging = false, offX = 0, offY = 0;
    const head = panel.querySelector('.head');
    head.addEventListener('mousedown', (e) => {
      if (e.target.closest('button')) return;
      dragging = true;
      const rect = panel.getBoundingClientRect();
      offX = e.clientX - rect.left; offY = e.clientY - rect.top;
      e.preventDefault();
    });
    window.addEventListener('mousemove', (e) => {
      if (!dragging) return;
      panel.style.left = (e.clientX - offX) + 'px';
      panel.style.top = (e.clientY - offY) + 'px';
      panel.style.right = 'auto';
    });
    window.addEventListener('mouseup', () => { dragging = false; });

    panel.querySelector('#ncMin').addEventListener('click', () => panel.classList.toggle('collapsed'));

    ui = {
      shadow, panel,
      excelFile: shadow.getElementById('ncExcelFile'),
      colRow: shadow.getElementById('ncColRow'),
      idCol: shadow.getElementById('ncIdCol'),
      scoreCol: shadow.getElementById('ncScoreCol'),
      excelStatus: shadow.getElementById('ncExcelStatus'),
      pdfFolder: shadow.getElementById('ncPdfFolder'),
      pdfStatus: shadow.getElementById('ncPdfStatus'),
      currentStat: shadow.getElementById('ncCurrentStat'),
      commentText: shadow.getElementById('ncCommentText'),
      fillGradeBtn: shadow.getElementById('ncFillGrade'),
      attachOnlyBtn: shadow.getElementById('ncAttachOnly'),
      doBothBtn: shadow.getElementById('ncDoBoth'),
      autoNext: shadow.getElementById('ncAutoNext'),
      autoRunBtn: shadow.getElementById('ncAutoRun'),
      log: shadow.getElementById('ncLog'),
    };
    return ui;
  }

  function addLog(id, name, action, ok, detail) {
    state.log.unshift({ time: fmtTime(), id: id || '-', name: name || '', action, ok, detail: detail || '' });
    renderLog();
  }

  function renderLog() {
    if (!state.log.length) { ui.log.innerHTML = '<div class="muted">尚無紀錄</div>'; return; }
    ui.log.innerHTML = state.log.slice(0, 60).map((l) => {
      const cls = l.ok === true ? 'ok' : l.ok === false ? 'warn' : 'muted';
      return `<div><span class="badge">${l.time}</span> <b>${l.id}</b> ${l.name || ''} — <span class="${cls}">${l.action}</span>${l.detail ? '：' + l.detail : ''}</div>`;
    }).join('');
  }

  function renderColumnPickers(headerRow) {
    ui.idCol.innerHTML = '';
    ui.scoreCol.innerHTML = '';
    headerRow.forEach((h, i) => {
      const label = (h === '' || h === undefined) ? `(第 ${i + 1} 欄)` : String(h);
      const o1 = document.createElement('option'); o1.value = i; o1.textContent = label; ui.idCol.appendChild(o1);
      const o2 = document.createElement('option'); o2.value = i; o2.textContent = label; ui.scoreCol.appendChild(o2);
    });
    const guess = guessColumns(headerRow);
    if (guess.idIdx >= 0) ui.idCol.value = guess.idIdx;
    if (guess.scoreIdx >= 0) ui.scoreCol.value = guess.scoreIdx;
    ui.colRow.style.display = 'flex';
  }

  function refreshCurrentStudentPanel() {
    const cur = getCurrentStudent();
    if (!cur) {
      ui.currentStat.innerHTML = '請開啟 SpeedGrader 並選擇一位學生';
      return null;
    }
    if (!cur.id) {
      ui.currentStat.innerHTML = `<span class="warn">無法從「${cur.raw}」解析出學號</span>`;
      return cur;
    }
    const score = state.scoreMap[cur.id];
    const pdf = state.pdfMap[cur.id];
    ui.currentStat.innerHTML = `
      <div><b>${cur.id}</b> ${cur.raw.replace(cur.id, '').trim()}</div>
      <div>成績比對：${score !== undefined ? `<span class="ok">${score}</span>` : '<span class="warn">找不到</span>'}</div>
      <div>PDF比對：${pdf ? `<span class="ok">${pdf.name}</span>` : '<span class="warn">找不到</span>'}</div>
    `;
    return cur;
  }

  // ---------------- 事件綁定 ----------------
  function wireEvents() {
    ui.excelFile.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      ui.excelStatus.textContent = '讀取中…';
      try {
        const { rows } = await parseWorkbookFile(file);
        if (!rows.length) throw new Error('檔案是空的');
        state.scoreRows = rows;
        renderColumnPickers(rows[0]);
        applyColumnSelection();
        ui.excelStatus.innerHTML = `<span class="ok">已載入 ${rows.length - 1} 筆資料（${file.name}）</span>`;
        saveScoreMapToStorage();
        refreshCurrentStudentPanel();
      } catch (err) {
        ui.excelStatus.innerHTML = `<span class="warn">讀取失敗：${err.message}</span>`;
      }
    });

    ui.idCol.addEventListener('change', applyColumnSelection);
    ui.scoreCol.addEventListener('change', applyColumnSelection);

    function applyColumnSelection() {
      state.idColIdx = Number(ui.idCol.value);
      state.scoreColIdx = Number(ui.scoreCol.value);
      state.scoreMap = buildScoreMap();
      const n = Object.keys(state.scoreMap).length;
      ui.excelStatus.innerHTML = `<span class="ok">已辨識 ${n} 位學生的成績</span>`;
      saveScoreMapToStorage();
      refreshCurrentStudentPanel();
    }

    ui.pdfFolder.addEventListener('change', (e) => {
      const files = e.target.files;
      if (!files || !files.length) return;
      const { map, unmatched } = indexPdfFiles(files);
      state.pdfMap = map;
      state.pdfUnmatched = unmatched;
      const n = Object.keys(map).length;
      ui.pdfStatus.innerHTML = `<span class="ok">已比對 ${n} 份 PDF</span>` +
        (unmatched.length ? `<br><span class="warn">${unmatched.length} 個檔名看不出學號</span>` : '');
      refreshCurrentStudentPanel();
    });

    ui.commentText.addEventListener('input', () => { state.commentText = ui.commentText.value; });
    ui.autoNext.addEventListener('change', () => { state.autoAdvance = ui.autoNext.checked; });

    ui.fillGradeBtn.addEventListener('click', () => runForCurrent({ grade: true, attach: false, submit: false }));
    ui.attachOnlyBtn.addEventListener('click', () => runForCurrent({ grade: false, attach: true, submit: false }));
    ui.doBothBtn.addEventListener('click', () => runForCurrent({ grade: true, attach: true, submit: true }));
    ui.autoRunBtn.addEventListener('click', runBatch);
  }

  // ---------------- 主要動作 ----------------
  async function runForCurrent(opts) {
    const cur = refreshCurrentStudentPanel();
    if (!cur || !cur.id) { addLog(cur && cur.raw, '', '略過', false, '無法辨識學號'); return false; }

    let ok = true;
    try {
      if (opts.grade) {
        const score = state.scoreMap[cur.id];
        if (score === undefined) {
          addLog(cur.id, '', '填成績', false, 'Excel 中找不到此學號');
          ok = false;
        } else {
          setGrade(score);
          addLog(cur.id, '', '填成績', true, score);
        }
      }
      if (opts.attach) {
        const file = state.pdfMap[cur.id];
        if (!file) {
          addLog(cur.id, '', '附加PDF', false, '資料夾中找不到此學號的PDF');
          ok = false;
        } else {
          await attachFile(file);
          addLog(cur.id, '', '附加PDF', true, file.name);
          if (state.commentText) setCommentText(state.commentText);
        }
      }
      if (opts.submit && ok) {
        const submitted = await submitComment();
        addLog(cur.id, '', '送出評論', submitted, submitted ? '' : '逾時，請手動確認畫面是否已送出');
        if (submitted && state.autoAdvance) {
          await wait(400);
          goNext();
        }
      }
    } catch (e) {
      addLog(cur.id, '', '錯誤', false, e.message);
      ok = false;
    }
    return ok;
  }

  async function runBatch() {
    const targetIds = new Set(Object.keys(state.scoreMap).filter((id) => state.pdfMap[id]));
    if (!targetIds.size) {
      alert('目前沒有「成績與PDF都比對成功」的學生可以批次處理。');
      return;
    }
    const confirmMsg = `即將依序處理最多 ${targetIds.size} 位「成績與PDF都已比對成功」的學生：\n填入成績 → 附加PDF → 送出評論 → 自動跳下一位。\n\n這會直接對 NTU COOL 上的真實成績與評論送出資料，且不易復原，請先確認已用「只填成績／只附加PDF」測試過至少一位學生沒問題。\n確定要開始嗎？`;
    if (!confirm(confirmMsg)) return;

    const sel = getStudentSelect();
    const rosterSize = sel ? sel.options.length : targetIds.size;
    const maxSteps = rosterSize + 5; // 安全上限，避免「下一位」在最後一位時繞回第一位造成無限迴圈
    const savedAutoAdvance = state.autoAdvance;

    ui.autoRunBtn.disabled = true;
    const processedIds = new Set();
    let processed = 0, failed = 0, steps = 0;
    try {
      while (processedIds.size < targetIds.size && steps < maxSteps) {
        steps++;
        const cur = getCurrentStudent();
        if (!cur || !cur.id) { addLog('-', '', '批次處理', false, '無法辨識目前學生，已中止'); break; }

        if (!targetIds.has(cur.id) || processedIds.has(cur.id)) {
          const moved = goNext();
          if (!moved) { addLog('-', '', '批次處理', false, '已到最後一位，停止批次'); break; }
          await wait(600);
          continue;
        }

        state.autoAdvance = true; // 批次處理時強制送出後自動換下一位，避免重複處理同一位
        const ok = await runForCurrent({ grade: true, attach: true, submit: true });
        processedIds.add(cur.id);
        if (ok) processed++; else failed++;
        if (!ok) {
          addLog(cur.id, '', '批次處理', false, '此位處理失敗，已停止批次（避免連續錯誤）');
          break;
        }
        await wait(800);
      }
      if (steps >= maxSteps) addLog('-', '', '批次處理', false, '已達安全上限次數，提前停止（請確認學生清單是否有循環）');
    } finally {
      state.autoAdvance = savedAutoAdvance;
      ui.autoRunBtn.disabled = false;
    }
    addLog('-', '', '批次處理完成', failed === 0, `成功 ${processed} 位，失敗 ${failed} 位`);
  }

  // ---------------- 監看學生切換 ----------------
  function watchStudentChange() {
    let lastValue = null;
    const check = () => {
      const cur = getCurrentStudent();
      const v = cur ? cur.optionValue : null;
      if (v !== lastValue) {
        lastValue = v;
        refreshCurrentStudentPanel();
      }
    };
    const sel = getStudentSelect();
    if (sel) sel.addEventListener('change', check);
    setInterval(check, 700); // 保險：輪詢比對，因為 Canvas 換頁不一定每次都會 fire change 事件
  }

  // ---------------- 初始化 ----------------
  async function init() {
    // 等待 SpeedGrader 主要元素出現
    for (let i = 0; i < 50; i++) {
      if (document.getElementById('grade_container')) break;
      await wait(200);
    }
    buildPanel();
    wireEvents();
    watchStudentChange();
    refreshCurrentStudentPanel();

    const saved = await loadScoreMapFromStorage();
    if (saved && Object.keys(saved).length) {
      state.scoreMap = saved;
      ui.excelStatus.innerHTML = `<span class="ok">已還原上次載入的 ${Object.keys(saved).length} 筆成績（PDF 資料夾請重新選擇）</span>`;
      refreshCurrentStudentPanel();
    }
  }

  init();
})();
