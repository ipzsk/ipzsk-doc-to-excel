import { readFile } from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import { pipeline } from 'node:stream/promises';
import { LargePDF, LARGE_SYSTEM } from './large.js';
import { Updater } from './updater.js';
const manifestText = await readFile(new URL('../package.json', import.meta.url),'utf8');
import { homedir } from 'node:os';
import path from 'node:path';
import { BlockAssembler, createUserMessage } from '@deepseek-ai/dsh-llm';
import { Engine, SYSTEM, PLAN_SYSTEM, findPython } from './engine.js';
export const name = 'doc-to-excel';
export const inject = ['webServer', 'llm'];
function guard(req) {
  if (!['127.0.0.1', '::1', '::ffff:127.0.0.1'].includes(req.socket.remoteAddress)) return false;
  try {
    const host = new URL(`http://${req.headers.host}`);
    return ['127.0.0.1', 'localhost', '[::1]'].includes(host.hostname) && req.headers['sec-fetch-site'] !== 'cross-site' && (!req.headers.origin || new URL(req.headers.origin).host === host.host);
  } catch { return false; }
}
async function body(req) {
  const chunks = []; let size = 0;
  for await (const chunk of req) { size += chunk.length; if (size > 29 * 1024 * 1024) throw new Error('上传超过 20 MB 限制'); chunks.push(chunk); }
  return JSON.parse(Buffer.concat(chunks).toString());
}
export function apply(ctx) {
  const profile = ctx.get('desktopProfiles')?.current?.dir;
  const directory = path.join(profile || path.join(homedir(), '.dsh'), 'doc-to-excel-jobs');
  const classify = async (model, batch, parentSignal) => {
    const signal = AbortSignal.any([parentSignal, AbortSignal.timeout(90000)]);
    const info = await ctx.llm.resolveModelInfo(model.provider, model.id, signal);
    const efforts = info.reasoning?.efforts ?? [];
    const effort = ['off', 'none', 'disabled', 'minimal', 'low'].find(id => efforts.some(e => e.id === id));
    const assembler = new BlockAssembler();
    for await (const chunk of ctx.llm.stream({ provider: model.provider, model: model.id, system: batch.task === 'large' ? LARGE_SYSTEM : batch.task === 'plan' ? PLAN_SYSTEM : SYSTEM, messages: [createUserMessage({content: [{type: 'text', text: JSON.stringify(batch)}], source: {kind: 'plugin', plugin: 'dsh-doc-to-excel'}})], maxTokens: batch.task === 'large' ? 8000 : batch.task === 'plan' ? 2200 : 3200, ...(effort ? {reasoningEffort: effort} : {}), signal })) { signal.throwIfAborted(); assembler.push(chunk); }
    if (assembler.finish?.kind !== 'stop') throw new Error(assembler.finish?.kind === 'max-tokens' ? '模型达到输出上限；请手动指定分类或选择推理开销较小的模型再试。' : '模型调用未正常完成；已停止后续调用。');
    return { text: assembler.blocks().filter(b => b.type === 'text').map(b => b.text).join(''), usage: assembler.usage };
  };
  const engine = new Engine({directory, classify});
  const large = new LargePDF({directory:path.join(directory,'large'), classify});
  const updater = new Updater({profileDir:profile,current:JSON.parse(manifestText).version});
  const routes = [
    ['GET', '/api/doc-to-excel/update/check', async () => updater.check()],
    ['GET', '/api/doc-to-excel/update/status', async () => updater.state],
    ['POST', '/api/doc-to-excel/update/install', async req => {
      if([...engine.jobs.values(),...large.jobs.values()].some(j=>j.running||j.busy)||engine.creating||large.starting)throw Error('请先暂停文档处理，再安装更新');
      return updater.install((await body(req)).version);
    }],
    ['GET', '/api/doc-to-excel/large/list', async () => large.list()],
    ['POST', '/api/doc-to-excel/large/init', async req => large.init(await body(req))],
    ['POST', '/api/doc-to-excel/large/upload', async req => {const u=new URL(req.url,'http://localhost');return large.upload(u.searchParams.get('id'),Number(u.searchParams.get('offset')),req);}],
    ['POST', '/api/doc-to-excel/large/start', async req => {if(updater.state.running)throw Error('更新期间请稍候');const {id,...options}=await body(req);return large.start(id,options);}],
    ['POST', '/api/doc-to-excel/large/pause', async req => large.pause((await body(req)).id)],
    ['POST', '/api/doc-to-excel/large/export', async req => large.export((await body(req)).id)],
    ['GET', '/api/doc-to-excel/large/status', async req => large.status(new URL(req.url,'http://localhost').searchParams.get('id'))],
    ['GET', '/api/doc-to-excel/large/download', async (req,res) => {
      const j=await large.get(new URL(req.url,'http://localhost').searchParams.get('id'));
      if(!j.download) throw Error('请先导出 Excel');
      res.writeHead(200,{'content-type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','content-disposition':'attachment; filename="pdf-records.xlsx"','cache-control':'no-store'});
      await pipeline(createReadStream(path.join(j.dir,'result.xlsx')),res);
    }],
    ['GET', '/api/doc-to-excel/config', async () => {
      let runtime = true, runtimeError = '';
      try { engine.executable ??= await findPython(); } catch (e) { runtime = false; runtimeError = e.message; }
      const models = [], errors = [];
      for (const p of ctx.llm.listProviders()) {
        try { for (const m of await ctx.llm.listModels(p.id)) models.push({provider: p.id, id: m.id, name: `${p.name} / ${m.name}`}); }
        catch { errors.push(p.name); }
      }
      return {runtime, runtimeError, models, modelWarnings: errors, providerCount: ctx.llm.listProviders().length};
    }],
    ['POST', '/api/doc-to-excel/create', async req => {if(updater.state.running)throw Error('更新期间请稍候');return engine.create(await body(req));}],
    ['GET', '/api/doc-to-excel/status', async req => engine.status(new URL(req.url, 'http://localhost').searchParams.get('id'))],
    ['POST', '/api/doc-to-excel/cancel', async req => { const {id} = await body(req); const j = engine.jobs.get(id); if (!j) throw new Error('任务不存在'); j.controller.abort(); return {ok: true}; }],
    ['GET', '/api/doc-to-excel/download', async (req, res) => {
      const id = new URL(req.url, 'http://localhost').searchParams.get('id'); const j = engine.jobs.get(id);
      if (!j?.download) throw new Error('Excel 尚未生成');
      const data = await readFile(path.join(j.dir, 'result.xlsx'));
      res.writeHead(200, {'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'content-disposition': 'attachment; filename="organized.xlsx"', 'cache-control': 'no-store'}); res.end(data);
    }],
  ];
  ctx.effect(() => {
    const disposers = routes.map(([method, routePath, handler]) => ctx.webServer.register({kind: 'exact', path: routePath, handler: async (req, res) => {
      const send = (status, data) => { res.writeHead(status, {'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store'}); res.end(JSON.stringify(data)); };
      if (!guard(req)) return send(403, {error: '仅支持本机同源访问'});
      if (req.method !== method) return send(405, {error: '请求方法不支持'});
      try { const result = await handler(req, res); if (!res.writableEnded) send(200, result); }
      catch (e) { if (!res.writableEnded) send(400, {error: String(e.message || e).slice(0, 500)}); }
    }}));
    return async () => { disposers.forEach(d => d()); await Promise.allSettled([engine.dispose(),large.dispose()]); };
  }, 'doc-to-excel routes and jobs');
}
