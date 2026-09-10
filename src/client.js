export const inject = [];
const API = '/api/doc-to-excel';
const LARGE_UI = `<div data-large>
<p>适合数千页 PDF。文件分块保存到本机，逐页解析并保存进度；模型仅接收文字片段及少量相邻上下文。</p>
<div class="dte-grid"><div>
<label>PDF 文件<input data-lg-file type="file" accept=".pdf"></label><p>单文件最多 2 GB；这是上传上限，复杂文件的处理速度取决于电脑和 OCR。当前支持本机运行。</p>
<label>筛选模型<input data-lg-model-filter placeholder="输入 qwen3.8 或提供方名称"></label><label>整理模型<select data-lg-model></select></label><p>优先选择已配置的 Qwen3.8；具体版本由你选择。</p>
<label>整理要求<textarea data-lg-requirements maxlength="2000" placeholder="例如：提取每家店铺的名称、地址、商品和金额。缺失留空，金额疑误标记待核对。"></textarea></label>
<label>Excel 列名（每行一列）<textarea data-lg-columns placeholder="店铺名称&#10;地址&#10;商品&#10;金额"></textarea></label><p>最多20列，留空输出「整理内容」。字段作为文本保存；本版不直接入库。跨页记录需核对合并。</p>
<label>本轮发送额度<select data-lg-budget><option value="48000">48,000 字符 · 试跑建议</option><option value="200000" selected>200,000 字符</option><option value="1000000">1,000,000 字符</option><option value="5000000">5,000,000 字符</option></select></label><p>含提示词与上下文；字符数非计费 token。每次继续授予新额度，用完暂停。每批输出最多8,000 token，不自动重试失败调用。</p>
<button data-lg-new class="dte-primary">上传并试跑 10 页</button>
<p>试跑会调用所选模型。修改模型、列名或要求后，创建新任务；继续会沿用原任务设置。</p>
</div><div>
<div class="dte-card"><label>已保存任务<select data-lg-history><option value="">选择任务</option></select></label><button data-lg-refresh>刷新任务</button>
<strong data-lg-phase style="display:block;margin-top:12px">准备就绪</strong><p data-lg-settings></p><p data-lg-stats class="dte-stats"></p><p data-lg-warning class="dte-warning"></p>
<button data-lg-continue disabled>继续全文</button> <button data-lg-pause hidden>暂停并保存</button> <button data-lg-retry disabled>重试解析失败页并继续</button>
<p><button data-lg-export disabled>导出当前结果</button> <a data-lg-download class="dte-download" hidden>下载 Excel</a></p>
<p>导出保留原文、未整理片段与异常页。关闭窗口可继续运行；退出桌面应用后需回来点继续。最近一次导出是快照，可暂停后重新导出。</p></div>
<div style="overflow:auto"><table><thead><tr><th>页码 / 片段</th><th>提取结果</th><th>原文</th></tr></thead><tbody data-lg-rows></tbody></table><p>预览前50个片段，完整结果在 Excel 中。</p></div>
</div></div></div>`;
function setupLarge(panel,request) {
  const q=s=>panel.querySelector(s); let id='',timer,disposed=false,busy=false,models=[];
  function renderModels(){const previous=q('[data-lg-model]').value;const term=q('[data-lg-model-filter]').value.toLowerCase();const list=models.filter(m=>(m.name+' '+m.id).toLowerCase().includes(term));q('[data-lg-model]').replaceChildren(new Option('选择模型',''));for(const m of list)q('[data-lg-model]').append(new Option(m.name,JSON.stringify(m)));const match=list.find(m=>/^qwen3[.]8-flash$/i.test(m.id))||list.find(m=>/qwen.?3[.]8/i.test(m.id));q('[data-lg-model]').value=list.some(m=>JSON.stringify(m)===previous)?previous:match?JSON.stringify(match):'';}
  q('[data-lg-model-filter]').oninput=renderModels;
  const budget=()=>Number(q('[data-lg-budget]').value);
  function error(e){q('[data-lg-warning]').textContent=e.message;}
  async function history(){const list=await request('/large/list');q('[data-lg-history]').replaceChildren(new Option('选择任务',''));for(const j of list)q('[data-lg-history]').append(new Option(j.name+' · '+j.phase,j.id));q('[data-lg-history]').value=id;}
  async function poll(){clearTimeout(timer);if(disposed||!id)return;
    try {const s=await request('/large/status?id='+encodeURIComponent(id));
      q('[data-lg-phase]').textContent=s.phase;
      q('[data-lg-settings]').textContent='任务列：'+s.columns.join('、')+'；模型：'+s.model.id+'；要求：'+(s.requirements||'按列提取');
      const usage=s.usage.reduce((a,u)=>{for(const [k,v] of Object.entries(u))if(typeof v==='number')a[k]=(a[k]||0)+v;return a;},{});
      q('[data-lg-stats]').textContent=`解析 ${s.extracted} / ${s.pages||'?'} 页 · 失败 ${s.failed} 页\n${s.total} 个文字片段 · ${s.pending} 个尚未整理\n累计 ${s.calls} 次模型调用 · ${s.payloadChars} 字符\n模型报告输入 ${usage.inputTokens||0} / 输出 ${usage.outputTokens||0} token（异常调用用量可能缺失）`;
      q('[data-lg-warning]').textContent=s.warning;
      q('[data-lg-continue]').disabled=s.running||s.uploaded!==s.size||s.complete;
      q('[data-lg-retry]').disabled=s.running||!s.failed;
      q('[data-lg-pause]').hidden=!s.running;
      q('[data-lg-export]').disabled=s.running||!s.extracted;
      q('[data-lg-new]').disabled=s.running||busy;
      const link=q('[data-lg-download]');link.hidden=!s.download;link.href=API+'/large/download?id='+encodeURIComponent(id);link.download='pdf-records.xlsx';
      q('[data-lg-rows]').replaceChildren();
      for(const r of s.preview){const tr=document.createElement('tr');for(const value of [r.page+' / '+r.id,r.decision.rows?JSON.stringify(r.decision.rows):'尚未整理',r.text]){const td=document.createElement('td');td.textContent=value;tr.append(td);}q('[data-lg-rows]').append(tr);}
      if(s.running)timer=setTimeout(poll,2500);
    }catch(e){error(e);if(!disposed)timer=setTimeout(poll,3000);}
  }
  async function start(options={}){try{await request('/large/start',{id,budget:budget(),...options});void poll();}catch(e){error(e);}}
  q('[data-lg-history]').onchange=()=>{id=q('[data-lg-history]').value;void poll();};
  q('[data-lg-refresh]').onclick=()=>{void history().catch(error);};
  q('[data-lg-continue]').onclick=()=>start();
  q('[data-lg-retry]').onclick=()=>start({retry:true});
  q('[data-lg-pause]').onclick=async()=>{try{await request('/large/pause',{id});void poll();}catch(e){error(e);}};
  q('[data-lg-export]').onclick=async()=>{q('[data-lg-export]').disabled=true;try{q('[data-lg-phase]').textContent='正在导出当前进度…';await request('/large/export',{id});await poll();}catch(e){error(e);q('[data-lg-export]').disabled=false;}};
  q('[data-lg-new]').onclick=async()=>{busy=true;q('[data-lg-new]').disabled=true;clearTimeout(timer);
    try {const file=q('[data-lg-file]').files[0];if(!file||!file.name.toLowerCase().endsWith('.pdf')||file.size>2*1024**3)throw Error('请选择不超过2GB的PDF');
      if(!q('[data-lg-model]').value)throw Error('请先选择模型');
      const result=await request('/large/init',{name:file.name,size:file.size,model:JSON.parse(q('[data-lg-model]').value),requirements:q('[data-lg-requirements]').value,columns:q('[data-lg-columns]').value});id=result.id;
      for(let offset=0;offset<file.size;offset+=4*1024**2){q('[data-lg-phase]').textContent=`上传到本机 ${Math.floor(offset/file.size*100)}%`;
        const r=await fetch(API+'/large/upload?id='+encodeURIComponent(id)+'&offset='+offset,{method:'POST',headers:{'Content-Type':'application/octet-stream'},body:file.slice(offset,offset+4*1024**2)});const v=await r.json();if(!r.ok)throw Error(v.error||'上传失败');}
      await history();await start({sample:true});
    }catch(e){error(e);}finally{busy=false;q('[data-lg-new]').disabled=false;}
  };
  return {async load(config){models=config.models;renderModels();await history();if(id)await poll();},dispose(){disposed=true;clearTimeout(timer);}};
}

export function apply(ctx) {
  ctx.effect(() => {
    const style = document.createElement('style');
    style.textContent = `
[data-dte-entry]{display:block;margin:6px 12px;padding:10px 14px;border:1px solid #68a99255;border-radius:9px;background:#19856816;color:inherit;cursor:pointer;text-align:left;font:inherit}
[data-dte-panel]{position:fixed;inset:65px 28px 24px 260px;z-index:100;background:light-dark(#f8faf9,#17221f);color:light-dark(#172b24,#e3eee8);border:1px solid #738c8044;border-radius:18px;box-shadow:0 18px 60px #0003;overflow:auto;padding:28px;font:14px/1.6 system-ui;box-sizing:border-box}
[data-dte-panel][hidden],[data-dte-panel] [hidden]{display:none!important}[data-dte-panel] h1{font-size:26px;margin:0}[data-dte-panel] p{color:light-dark(#5b6e64,#9db7aa)}
[data-dte-panel] .dte-grid{display:grid;grid-template-columns:minmax(240px,310px) minmax(260px,1fr);gap:24px;margin-top:24px}[data-dte-panel] label{display:block;margin:0 0 16px;font-weight:600}
[data-dte-panel] select,[data-dte-panel] textarea,[data-dte-panel] input{display:block;box-sizing:border-box;width:100%;margin-top:7px;padding:10px;border:1px solid #78988855;border-radius:8px;background:light-dark(white,#22312b);color:inherit;font:inherit}
[data-dte-panel] textarea{min-height:135px;resize:vertical}[data-dte-panel] button,[data-dte-panel] a.dte-download{padding:9px 16px;border:1px solid #73908066;border-radius:8px;background:light-dark(white,#22312b);color:inherit;cursor:pointer;font:inherit;text-decoration:none;display:inline-block}
[data-dte-panel] button:disabled{opacity:.5;cursor:wait}[data-dte-panel] .dte-primary{background:#16815e;color:white}[data-dte-panel] .dte-card{padding:20px;border:1px solid #78988844;border-radius:12px;background:light-dark(white,#1c2b24)}
[data-dte-panel] table{width:100%;border-collapse:collapse;font-size:12px}[data-dte-panel] td,[data-dte-panel] th{text-align:left;padding:8px;border-bottom:1px solid #78988833;vertical-align:top;word-break:break-word}[data-dte-panel] .dte-stats{white-space:pre-wrap}[data-dte-panel] .dte-warning{color:light-dark(#a34e00,#eab77a)}
@media(max-width:850px){[data-dte-panel]{left:16px;right:16px}[data-dte-panel] .dte-grid{grid-template-columns:1fr}}`;
    document.head.append(style);
    const panel = document.createElement('section'); panel.dataset.dtePanel = ''; panel.hidden = true; panel.setAttribute('aria-label', '文档转 Excel');
    panel.innerHTML = `<div style="display:flex;justify-content:space-between;align-items:center"><div><h1>文档转 Excel</h1><p>把散落的文档，整理成有分类、有出处的表格。</p></div><button data-close aria-label="关闭文档面板">关闭</button></div>
<div class="dte-card"><button data-update-check>检查插件更新</button> <button data-update-install hidden>安装更新</button><p data-update-status>独立更新插件，处理数据保存在本机。</p><a data-update-release target="_blank" rel="noopener" hidden>查看发布说明</a></div><label>处理模式<select data-mode><option value="large">大 PDF · 按列提取与断点续跑</option><option value="small">小文档 · 分类整理（原有功能）</option></select></label>${LARGE_UI}<div data-small hidden><div class="dte-grid"><div><label>上传文档<input data-files type="file" multiple accept=".docx,.pdf"></label><p>支持 Word .docx、PDF；最多 10 个文件，合计 20 MB。macOS 可本地识别正文图片与扫描页，OCR 结果标记待确认。</p><label>分类模型<select data-model><option value="">纯本地规则 · 不调用模型</option></select></label><label>整理要求<textarea data-requirements maxlength="2000" placeholder="像聊天一样告诉我：要提取什么、如何分类、哪些不要。&#10;例如：按商家主体、商品信息、宣传承诺、证据来源分类，去掉广告，金额和编号不确定时标记待核对。"></textarea></label><p>填写要求后需选择模型。未指定分类时会根据要求设计分类；修改要求后可重新整理。</p><label>指定分类（可选，每行一个）<textarea data-categories placeholder="有整理要求时留空自动设计分类&#10;需求：需求,问题&#10;计划：计划,截止&#10;数据：金额,收入"></textarea></label><label>模型摘录预算<select data-budget><option value="6000">6,000 字符 · 节省</option><option value="12000" selected>12,000 字符 · 标准</option><option value="24000">24,000 字符</option><option value="48000">48,000 字符</option></select></label><p>仅发送待分类摘录，不携带聊天历史。字符预算不等于计费 token；模型调用最多 12 次，每批输出上限 3,200 token；设计分类最多额外占其中 1 次调用。</p><button data-start class="dte-primary" disabled>开始整理</button> <button data-cancel hidden>停止</button></div>
<div><div class="dte-card"><strong data-phase>正在检查运行环境…</strong><p data-stats class="dte-stats">选择文件后开始整理。</p><p data-warning class="dte-warning"></p><a data-download class="dte-download" hidden>下载 Excel</a></div><div style="overflow:auto;margin-top:18px"><table><thead><tr><th>分类</th><th>内容预览</th><th>来源</th></tr></thead><tbody data-rows></tbody></table><p>预览显示前 100 条，每条最多 300 字，点击可核对原始提取。Excel 同时保留原文及被排除内容。</p></div></div></div></div>`;
    document.body.append(panel);
    const q = s => panel.querySelector(s), entry = document.createElement('button'); entry.dataset.dteEntry = ''; entry.textContent = '▦  文档转 Excel';
    let timer, disposed = false, current, config, loading = false;
    async function request(url, data) { const r = await fetch(API + url, data === undefined ? {} : {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}); const v = await r.json(); if (!r.ok) throw new Error(v.error || '请求失败'); return v; }
    let updateVersion,updateTimer;
    async function pollUpdate(){if(disposed)return;try{const s=await request('/update/status');q('[data-update-status]').textContent=s.phase;if(s.running)updateTimer=setTimeout(pollUpdate,1500);else q('[data-update-check]').disabled=false;}catch(e){q('[data-update-status]').textContent=e.message+'；可重启客户端检查版本';q('[data-update-check]').disabled=false;}}
    q('[data-update-check]').onclick=async()=>{q('[data-update-check]').disabled=true;try{const r=await request('/update/check');updateVersion=r.version;q('[data-update-status]').textContent=r.available?`当前 ${r.current}，发现 ${r.version}；请先暂停任务再更新`:`当前 ${r.current}，已是最新版本`;q('[data-update-install]').hidden=!r.available||!r.canInstall;q('[data-update-install]').disabled=false;const link=q('[data-update-release]');link.hidden=false;link.href=r.releaseURL;}catch(e){q('[data-update-status]').textContent=e.message;}finally{q('[data-update-check]').disabled=false;}};
    q('[data-update-install]').onclick=async()=>{q('[data-update-install]').disabled=true;try{await request('/update/install',{version:updateVersion});q('[data-update-check]').disabled=true;void pollUpdate();}catch(e){q('[data-update-status]').textContent=e.message;q('[data-update-install]').disabled=false;}};
    const largeUI=setupLarge(panel,request);
    q('[data-mode]').onchange=()=>{q('[data-large]').hidden=q('[data-mode]').value!=='large';q('[data-small]').hidden=q('[data-mode]').value!=='small';};
    async function load() {
      if (loading) return; loading = true;
      try { const previousModel = q('[data-model]').value; config = await request('/config'); q('[data-model]').replaceChildren(new Option('纯本地规则 · 不调用模型', '')); for (const m of config.models) { const option = document.createElement('option'); option.value = JSON.stringify(m); option.textContent = m.name; q('[data-model]').append(option); }
        if ([...q('[data-model]').options].some(o => o.value === previousModel)) q('[data-model]').value = previousModel;
        q('[data-phase]').textContent = config.runtime ? '准备就绪' : '需要配置运行环境'; q('[data-warning]').textContent = config.runtimeError || (config.models.length ? '' : `暂无可用模型（${config.providerCount} 个提供方）；${config.modelWarnings.join('、')}。请检查 Harness 模型配置。`); q('[data-start]').disabled = !config.runtime; await largeUI.load(config);
      } catch (e) { q('[data-phase]').textContent = '加载失败'; q('[data-warning]').textContent = e.message; } finally { loading = false; }
    }
    const activate = e => { if (e.detail !== 'doc-to-excel') panel.hidden = true; };
    entry.onclick = () => { window.dispatchEvent(new CustomEvent('dsh-panel-activate', {detail:'doc-to-excel'})); panel.hidden = false; void load(); };
    q('[data-close]').onclick = () => { panel.hidden = true; };
    window.addEventListener('dsh-panel-activate', activate);
    function mount() {
      const anchor = document.querySelector('[data-dsh-imagegen-session-tabs],button[data-dsh-part="new-session"],button[class*="newSession"]');
      if (anchor) {
        if (entry.previousElementSibling !== anchor) anchor.after(entry);
        entry.style.cssText = '';
      } else if (!entry.isConnected) {
        document.body.append(entry);
        entry.style.cssText = 'position:fixed;left:16px;bottom:62px;z-index:90;background:var(--dsw-alias-bg-base,#183228)';
      }
    }
    const observer = new MutationObserver(mount); observer.observe(document.body, {subtree:true,childList:true}); mount();
    async function poll() {
      if (disposed || !current) return;
      try { const s = await request('/status?id=' + encodeURIComponent(current)); q('[data-phase]').textContent = s.cached ? '已复用整理结果 · 本次未调用模型' : s.phase;
        const usage = s.usage.reduce((a,u) => { for (const k of ['inputTokens','outputTokens','cacheReadTokens','cacheWriteTokens']) a[k] = (a[k] || 0) + (Number(u[k]) || 0); return a; }, {});
        q('[data-stats]').textContent = `${s.total} 条内容 · ${s.review} 条待确认 · ${s.excluded || 0} 条已排除\n模型调用 ${s.calls} 次 · 发送 ${s.payloadChars} 字符（含分类提示）\n` + (s.usage.length ? `已报告用量：输入 ${usage.inputTokens || 0} / 缓存读取 ${usage.cacheReadTokens || 0} / 缓存写入 ${usage.cacheWriteTokens || 0} / 输出 ${usage.outputTokens || 0} token\n` : '') + Object.entries(s.counts).map(([k,v]) => `${k} ${v}`).join(' · ');
        q('[data-warning]').textContent = s.warning; q('[data-rows]').replaceChildren();
        for (const row of s.preview) { const tr = document.createElement('tr'); for (const text of [(row.excluded ? '已排除 · ' : '') + row.category,row.text,row.source+' · '+row.location]) { const td=document.createElement('td');td.textContent=text;tr.append(td); } const details=document.createElement('details');const summary=document.createElement('summary');summary.textContent='核对原始提取与分类理由';const raw=document.createElement('p');raw.textContent=row.original+'\n理由：'+row.reason;details.append(summary,raw);tr.children[1].append(details);q('[data-rows]').append(tr); }
        const download = q('[data-download]'); download.hidden = !s.download; download.href = API + '/download?id=' + encodeURIComponent(current); download.setAttribute('download','organized.xlsx');
        q('[data-cancel]').hidden = !s.running; q('[data-start]').disabled = s.running; q('[data-start]').textContent = s.running ? '正在整理…' : '按当前要求重新整理';
        if (s.running) timer = setTimeout(poll, 1800);
      } catch (e) { q('[data-warning]').textContent = e.message; q('[data-start]').disabled = false; }
    }
    q('[data-start]').onclick = async () => {
      q('[data-start]').disabled = true; clearTimeout(timer); q('[data-download]').hidden = true;
      try { const selected = [...q('[data-files]').files]; if (!selected.length || selected.length > 10 || selected.reduce((n,f)=>n+f.size,0)>20*1024*1024) throw new Error('请选择 1–10 个文件，合计不超过 20 MB');
        const files = await Promise.all(selected.map(f => new Promise((resolve,reject) => { const reader = new FileReader(); reader.onload = () => resolve({name:f.name,data:String(reader.result).split(',')[1]}); reader.onerror = () => reject(new Error('文件读取失败')); reader.readAsDataURL(f); })));
        const modelValue = q('[data-model]').value; const r = await request('/create', {files,model:modelValue ? JSON.parse(modelValue) : null,budget:Number(q('[data-budget]').value),categories:q('[data-categories]').value,requirements:q('[data-requirements]').value}); current = r.id; void poll();
      } catch (e) { q('[data-warning]').textContent = e.message; q('[data-start]').disabled = false; }
    };
    q('[data-cancel]').onclick = async () => { try { await request('/cancel',{id:current}); } catch (e) { q('[data-warning]').textContent=e.message; } };
    return () => { disposed=true; largeUI.dispose(); clearTimeout(timer); clearTimeout(updateTimer); observer.disconnect(); window.removeEventListener('dsh-panel-activate', activate); entry.remove(); panel.remove(); style.remove(); };
  }, 'doc-to-excel panel');
}
