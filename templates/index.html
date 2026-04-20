<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>協力会社フォーム自動送信システム</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; background: #f0f4f8; color: #333; }
    .header { background: #1F4E79; color: white; padding: 20px 32px; }
    .header h1 { font-size: 20px; }
    .header p { font-size: 13px; opacity: 0.8; margin-top: 4px; }
    .container { max-width: 900px; margin: 32px auto; padding: 0 16px; }
    .card { background: white; border-radius: 12px; padding: 28px; margin-bottom: 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .card h2 { font-size: 16px; color: #1F4E79; margin-bottom: 16px; border-left: 4px solid #1F4E79; padding-left: 10px; }
    .step-badge { display: inline-block; background: #1F4E79; color: white; border-radius: 50%; width: 24px; height: 24px; text-align: center; line-height: 24px; font-size: 12px; margin-right: 8px; }
    .upload-area { border: 2px dashed #aac4e0; border-radius: 8px; padding: 40px; text-align: center; cursor: pointer; transition: all 0.2s; }
    .upload-area:hover { border-color: #1F4E79; background: #f0f7ff; }
    .upload-area input { display: none; }
    .upload-area p { color: #666; font-size: 14px; margin-top: 8px; }
    .preview-box { background: #f8fafb; border-radius: 8px; padding: 16px; margin-top: 16px; font-size: 13px; }
    .preview-box table { width: 100%; border-collapse: collapse; }
    .preview-box th { background: #1F4E79; color: white; padding: 8px 12px; text-align: left; }
    .preview-box td { padding: 8px 12px; border-bottom: 1px solid #eee; }
    .sender-info { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 12px; }
    .sender-item { background: #f0f7ff; border-radius: 6px; padding: 10px 14px; }
    .sender-item label { font-size: 11px; color: #666; display: block; }
    .sender-item span { font-size: 13px; font-weight: bold; }
    .mode-selector { display: flex; gap: 12px; margin-bottom: 16px; }
    .mode-btn { flex: 1; padding: 14px; border: 2px solid #ddd; border-radius: 8px; background: white; cursor: pointer; text-align: center; transition: all 0.2s; }
    .mode-btn.active { border-color: #1F4E79; background: #f0f7ff; }
    .mode-btn h3 { font-size: 14px; margin-bottom: 4px; }
    .mode-btn p { font-size: 12px; color: #666; }
    .mode-btn.production.active { border-color: #C00000; background: #fff0f0; }
    .run-btn { width: 100%; padding: 16px; background: #1F4E79; color: white; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; }
    .run-btn:disabled { background: #aaa; cursor: not-allowed; }
    .run-btn.production { background: #C00000; }
    .progress-bar-wrap { background: #eee; border-radius: 99px; height: 12px; margin: 12px 0; }
    .progress-bar { background: #1F4E79; height: 12px; border-radius: 99px; transition: width 0.3s; width: 0%; }
    .progress-text { font-size: 13px; color: #555; text-align: center; }
    .log-box { background: #1a1a2e; color: #7ec8e3; border-radius: 8px; padding: 16px; max-height: 300px; overflow-y: auto; font-size: 12px; font-family: monospace; margin-top: 12px; }
    .log-box .ok { color: #69db7c; }
    .log-box .err { color: #ff6b6b; }
    .log-box .skip { color: #ffd43b; }
    .download-btn { display: none; width: 100%; padding: 14px; background: #375623; color: white; border: none; border-radius: 8px; font-size: 15px; font-weight: bold; cursor: pointer; margin-top: 16px; }
    .warning-box { background: #fff8e6; border: 1px solid #ffc107; border-radius: 8px; padding: 14px 16px; font-size: 13px; color: #856404; margin-bottom: 16px; }
  </style>
</head>
<body>
<div class="header">
  <h1>🤖 協力会社フォーム自動送信システム</h1>
  <p>Excelをアップロードするだけで、問い合わせフォームへの自動送信が完了します</p>
</div>
<div class="container">
  <div class="card">
    <h2><span class="step-badge">1</span>Excelファイルをアップロード</h2>
    <div class="upload-area" onclick="document.getElementById('fileInput').click()">
      <input type="file" id="fileInput" accept=".xlsx" onchange="uploadFile(this)">
      <div style="font-size:40px">📂</div>
      <p>クリックしてExcelファイルを選択<br><small>（協力会社リスト.xlsx）</small></p>
    </div>
    <div id="previewArea" style="display:none">
      <div style="margin-top:16px;font-size:14px;font-weight:bold;color:#1F4E79">📋 送信者情報</div>
      <div class="sender-info" id="senderInfo"></div>
      <div style="margin-top:16px;font-size:14px;font-weight:bold;color:#1F4E79">🏢 協力会社リスト（<span id="companyCount">0</span>社）</div>
      <div class="preview-box">
        <table><thead><tr><th>会社名</th><th>URL</th></tr></thead><tbody id="previewBody"></tbody></table>
      </div>
    </div>
  </div>

  <div class="card" id="step2" style="display:none">
    <h2><span class="step-badge">2</span>実行モードを選択</h2>
    <div class="warning-box">⚠️ 初めて使う場合は必ず<strong>テストモード</strong>で動作確認してください</div>
    <div class="mode-selector">
      <div class="mode-btn active" id="testBtn" onclick="selectMode('test')">
        <h3>🧪 テストモード</h3>
        <p>フォームを探して入力内容を確認するだけ。実際には送信しません</p>
      </div>
      <div class="mode-btn production" id="prodBtn" onclick="selectMode('production')">
        <h3>🚀 本番モード</h3>
        <p>実際にフォームを送信します。取り消しできません</p>
      </div>
    </div>
    <button class="run-btn" id="runBtn" onclick="startRun()">🧪 テストモードで実行する</button>
  </div>

  <div class="card" id="step3" style="display:none">
    <h2><span class="step-badge">3</span>実行中...</h2>
    <div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div>
    <div class="progress-text" id="progressText">準備中...</div>
    <div class="log-box" id="logBox"></div>
    <button class="download-btn" id="downloadBtn" onclick="downloadResult()">📥 結果Excelをダウンロード</button>
  </div>
</div>

<script>
let currentMode = 'test', currentSessionId = null, progressInterval = null;

function selectMode(mode) {
  currentMode = mode;
  document.getElementById('testBtn').classList.toggle('active', mode === 'test');
  document.getElementById('prodBtn').classList.toggle('active', mode === 'production');
  const btn = document.getElementById('runBtn');
  btn.className = mode === 'production' ? 'run-btn production' : 'run-btn';
  btn.textContent = mode === 'test' ? '🧪 テストモードで実行する' : '🚀 本番モードで送信する';
}

async function uploadFile(input) {
  const file = input.files[0];
  if (!file) return;
  const form = new FormData();
  form.append('file', file);
  const res = await fetch('/upload', { method: 'POST', body: form });
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  const s = data.sender;
  document.getElementById('senderInfo').innerHTML = `
    <div class="sender-item"><label>送信者名</label><span>${s['送信者名']||'-'}</span></div>
    <div class="sender-item"><label>会社名</label><span>${s['送信者会社名']||'-'}</span></div>
    <div class="sender-item"><label>メール</label><span>${s['メール']||'-'}</span></div>
    <div class="sender-item"><label>電話</label><span>${s['電話']||'-'}</span></div>`;
  document.getElementById('companyCount').textContent = data.count;
  const tbody = document.getElementById('previewBody');
  tbody.innerHTML = '';
  data.companies.slice(0,10).forEach(c => {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${c['会社名']}</td><td><a href="${c['URL']}" target="_blank">${c['URL']}</a></td>`;
    tbody.appendChild(tr);
  });
  if (data.count > 10) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="2" style="color:#999;text-align:center">... 他${data.count-10}社</td>`;
    tbody.appendChild(tr);
  }
  document.getElementById('previewArea').style.display = 'block';
  document.getElementById('step2').style.display = 'block';
}

async function startRun() {
  if (currentMode === 'production' && !confirm('本番モードで実際に送信します。よろしいですか？')) return;
  document.getElementById('runBtn').disabled = true;
  document.getElementById('step3').style.display = 'block';
  document.getElementById('step3').scrollIntoView({behavior:'smooth'});
  const res = await fetch('/run', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode:currentMode})});
  const data = await res.json();
  currentSessionId = data.session_id;
  progressInterval = setInterval(checkProgress, 1500);
}

async function checkProgress() {
  const res = await fetch(`/progress/${currentSessionId}`);
  const data = await res.json();
  const pct = data.total > 0 ? Math.round(data.done/data.total*100) : 0;
  document.getElementById('progressBar').style.width = pct + '%';
  document.getElementById('progressText').textContent = `${data.done} / ${data.total} 社処理済み（${pct}%）`;
  const logBox = document.getElementById('logBox');
  logBox.innerHTML = '';
  (data.results||[]).slice(-20).reverse().forEach(r => {
    const cls = r.status.includes('完了') ? 'ok' : r.status.includes('スキップ') ? 'skip' : 'err';
    logBox.innerHTML += `<div class="${cls}">${r.status} | ${r.company} | ${r.reason||r.form_url||''}</div>`;
  });
  if (data.status === 'done') {
    clearInterval(progressInterval);
    document.getElementById('progressText').textContent = `✅ 完了！ ${data.total}社の処理が終わりました`;
    document.getElementById('downloadBtn').style.display = 'block';
    document.getElementById('step3').querySelector('h2').textContent = '✅ 完了';
  } else if (data.status === 'error') {
    clearInterval(progressInterval);
    document.getElementById('progressText').textContent = `❌ エラー: ${data.error}`;
  }
}
function downloadResult() { window.location.href = `/download/${currentSessionId}`; }
</script>
</body>
</html>
