import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, readFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { Engine, parseDecisions, findPython, runProcess } from '../lib/engine.js';
const fixtureRoot = await mkdtemp(path.join(os.tmpdir(), 'dsh-doc-excel-test-'));
const python = await findPython();
await runProcess(python, ['-c', 'from docx import Document;import sys;d=Document();d.add_paragraph("客户需求：离线使用");d.add_paragraph("清晨空气很好");d.add_paragraph("清晨空气很好");d.save(sys.argv[1])', path.join(fixtureRoot, 'demo.docx')]);
const files = [{name:'demo.docx',data:(await readFile(path.join(fixtureRoot,'demo.docx'))).toString('base64')}];
const batch = {categories:['其他'],records:[{id:'a'}]};
test('model response rejects missing, extra, duplicate, invalid IDs and NaN', () => {
  for (const x of ['[]','[{"id":"wrong","category":"其他","confidence":1,"reason":""}]','[{"id":"a","category":"其他","confidence":2,"reason":""}]']) assert.throws(()=>parseDecisions(x,batch));
  assert.equal(parseDecisions('```json\n[{"id":"a","category":"其他","confidence":0.5,"reason":""}]\n```',batch).length,1);
});
test('local export performs no LLM calls and preserves unresolved rows',async()=>{
 const e = new Engine({directory:path.join(fixtureRoot,'local'),classify:()=>{throw Error('Must not call');}});
 const {id}=await e.create({files});await e.jobs.get(id).done;
 const s=await e.status(id); assert.equal(s.download,true);assert.equal(s.calls,0);assert.equal(s.total,3);assert.equal(s.review,2);
});
test('bounded one-shot classification deduplicates and never carries chat history',async()=>{
 let calls=0;
 const e=new Engine({directory:path.join(fixtureRoot,'model'),classify:async(model,b)=>{calls++;assert.equal(b.records.length,1);return {text:JSON.stringify(b.records.map(r=>({id:r.id,category:'其他',confidence:.9,reason:'示例'}))),usage:{inputTokens:120,outputTokens:30}};}});
 const {id}=await e.create({files,model:{provider:'mock',id:'mock'}});await e.jobs.get(id).done;
 const s=await e.status(id);assert.equal(calls,1);assert.equal(s.total,3);assert.equal(s.review,0);assert.equal(s.download,true);
});
test('malformed model output stops without retries and still exports',async()=>{
 let calls=0;const e=new Engine({directory:path.join(fixtureRoot,'bad'),classify:async()=>{calls++;return {text:'not json'};}});
 const {id}=await e.create({files,model:{provider:'mock',id:'mock'}});await e.jobs.get(id).done;
 const s=await e.status(id);assert.equal(calls,1);assert.equal(s.download,true);assert.equal(s.review,2);assert.ok(s.warning);
});
test('cancel interrupts model and exports local progress',async()=>{
 let entered;const started=new Promise(r=>entered=r);
 const e=new Engine({directory:path.join(fixtureRoot,'cancel'),classify:async(m,b,signal)=>{entered();return new Promise((resolve,reject)=>{signal.addEventListener('abort',()=>reject(Error('cancelled')),{once:true});});}});
 const {id}=await e.create({files,model:{provider:'mock',id:'mock'}});await started;e.jobs.get(id).controller.abort();await e.jobs.get(id).done;
 const s=await e.status(id);assert.equal(s.download,true);assert.equal(s.running,false);assert.match(s.warning,/已停止/);
});
test('completed same-input job reuses disk cache across plugin restarts',async()=>{
 const directory=path.join(fixtureRoot,'cache');let calls=0;
 const classify=async(m,b)=>{calls++;return {text:JSON.stringify(b.records.map(r=>({id:r.id,category:'其他',confidence:.9,reason:'示例'})))}};
 const input={files,model:{provider:'mock',id:'mock'}};
 const e=new Engine({directory,classify});const {id}=await e.create(input);await e.jobs.get(id).done;
 const restarted=new Engine({directory,classify});const r=await restarted.create(input);
 assert.equal(r.id,id);assert.equal(calls,1);assert.equal((await restarted.status(id)).cached,true);
});
test('natural-language requirements design categories and override keyword rules',async()=>{
 let calls=0;let batchSeen;
 const e=new Engine({directory:path.join(fixtureRoot,'guided'),classify:async(m,b)=>{
 calls++;
 if(b.task==='plan') return {text:JSON.stringify({categories:['客户事项','无关内容']})};
 batchSeen=b;
 return {text:JSON.stringify(b.records.map(r=>({id:r.id,category:r.excerpt.includes('客户')?'客户事项':'无关内容',confidence:.9,reason:'按用户要求',organized_text:r.excerpt.includes('客户')?'客户要求支持离线使用':'无关描写',include:r.excerpt.includes('客户')})))};
 }});
 const input={files,requirements:'仅保留客户要求，排除景色描述',model:{provider:'mock',id:'mock'}};
 const {id}=await e.create(input);await e.jobs.get(id).done;
 const s=await e.status(id);assert.equal(s.download,true);assert.equal(calls,2);assert.equal(batchSeen.records.length,2);assert.equal(batchSeen.userRequirements,input.requirements);assert.equal(s.excluded,2);assert.equal(s.preview[0].text,'客户要求支持离线使用');
 const report=await runProcess(python,['-c','import openpyxl,sys,json;w=openpyxl.load_workbook(sys.argv[1]);print(json.dumps([w["分类明细"].max_row,w["已排除"].max_row,w["分类明细"]["J2"].value],ensure_ascii=False))',path.join(e.jobs.get(id).dir,'result.xlsx')]);
 assert.deepEqual(JSON.parse(report),[2,3,'客户需求：离线使用']);
 const next=await e.create({...input,requirements:'保留全部，区分客户事项和无关内容'});assert.notEqual(next.id,id);await e.jobs.get(next.id).done;
});
test('requirements cannot silently run in local mode and invalid plans are rejected',async()=>{
 const e=new Engine({directory:path.join(fixtureRoot,'invalid-guided'),classify:()=>{throw Error('not called')}});
 await assert.rejects(e.create({files,requirements:'按主体分类'}),/选择一个模型/);
});
test('guided long content is split without sending truncated record prefixes',async()=>{
 const file=path.join(fixtureRoot,'long.docx');await runProcess(python,['-c','from docx import Document;import sys;d=Document();d.add_paragraph("甲乙丙丁"*750);d.save(sys.argv[1])',file]);
 const e=new Engine({directory:path.join(fixtureRoot,'long-guided'),classify:async(m,b)=>{assert.ok(b.records.every(r=>!r.truncated&&r.excerpt.length<=1000));return {text:JSON.stringify(b.records.map(r=>({id:r.id,category:'保留',confidence:.5,reason:'待核对',organized_text:r.excerpt,include:true})))}}});
 const {id}=await e.create({files:[{name:'long.docx',data:(await readFile(file)).toString('base64')}],model:{provider:'mock',id:'mock'},requirements:'保留原文',categories:'保留\n其他'});await e.jobs.get(id).done;assert.equal((await e.status(id)).download,true);
 const state=JSON.parse(await readFile(path.join(e.jobs.get(id).work,'state.json'),'utf8'));assert.equal(state.records.map(r=>r.text).join(''),'甲乙丙丁'.repeat(750));
});
test('valid JSON with explanatory prose is parsed locally without another model call',()=>{
 const result=parseDecisions('分类如下：\n[{"id":"a","category":"其他","confidence":0.5,"reason":"原文含方括号 [ ]"}]\n以上分类供核对。',batch);assert.equal(result.length,1);
 assert.throws(()=>parseDecisions('[{"id":"a","category":"其他","confidence":0.5,"reason":""}',batch));
 assert.throws(()=>parseDecisions('[]\n[]',batch));
});
