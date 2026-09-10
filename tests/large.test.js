import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,readFile,writeFile,stat} from 'node:fs/promises';
import {Readable} from 'node:stream';
import path from 'node:path';
import os from 'node:os';
import {LargePDF,validateRows} from '../lib/large.js';
import {findPython,runProcess} from '../lib/engine.js';
const root=await mkdtemp(path.join(os.tmpdir(),'dte-large-'));const python=await findPython();
await runProcess(python,['-c',`from reportlab.pdfgen import canvas
import sys
c=canvas.Canvas(sys.argv[1])
for p in range(24):
 c.drawString(50,700, 'Shop %s amount 123.45' % (p+1));c.showPage()
c.save()`,path.join(root,'sample.pdf')]);
const bytes=await readFile(path.join(root,'sample.pdf'));
async function upload(e,model={provider:'mock',id:'qwen3.8'}){const {id}=await e.init({name:'sample.pdf',size:bytes.length,columns:'商家\n金额',requirements:'提取每家商店',model});for(let n=0;n<bytes.length;n+=1024)await e.upload(id,n,Readable.from([bytes.subarray(n,n+1024)]));return id;}
const response=b=>({text:JSON.stringify(b.records.map(r=>({id:r.id,rows:[{values:{'商家':'=literal','金额':'123.45'},review:''}],note:''}))),usage:{inputTokens:100,outputTokens:50}});
test('large PDF: sample, persistent restart, resume, export literal fields and all source pages',async()=>{
 let calls=0;const classify=async(m,b)=>{calls++;assert.equal(m.id,'qwen3.8');assert.deepEqual(b.columns,['商家','金额']);return response(b);};
 const directory=path.join(root,'resume'),e=new LargePDF({directory,classify});const id=await upload(e);await e.start(id,{sample:true});await e.jobs.get(id).done;
 let s=await e.status(id);assert.equal(s.extracted,10);assert.equal(s.pending,0);assert.equal(s.complete,false);assert.equal(calls,5);
 const next=new LargePDF({directory,classify});assert.equal((await next.list()).length,1);await next.start(id);await next.jobs.get(id).done;s=await next.status(id);
 assert.equal(s.pages,24);assert.equal(s.complete,true);assert.equal(s.download,true);assert.equal(s.calls,12);assert.equal(calls,12);
 const check=JSON.parse(await runProcess(python,['-c',`import openpyxl,json,sys
w=openpyxl.load_workbook(sys.argv[1],read_only=True)
r=list(w['提取结果'].rows)
print(json.dumps([len(r),r[1][1].value,r[1][1].data_type,len(list(w['原文'].rows))]))`,path.join(next.jobs.get(id).dir,'result.xlsx')]));
 assert.deepEqual(check,[25,'=literal','s',25]);
});
test('budget pause never claims complete and explicit resume handles remaining records',async()=>{
 const e=new LargePDF({directory:path.join(root,'budget'),classify:async(m,b)=>response(b)});const id=await upload(e);await e.start(id,{budget:6000});await e.jobs.get(id).done;
 let s=await e.status(id);assert.equal(s.extracted,24);assert.equal(s.complete,false);assert.ok(s.pending>0);assert.match(s.warning,/额度/);
 await e.start(id);await e.jobs.get(id).done;assert.equal((await e.status(id)).complete,true);
});
test('bad output pauses once; retry occurs only on explicit continuation',async()=>{
 let calls=0;const e=new LargePDF({directory:path.join(root,'bad'),classify:async(m,b)=>{calls++;return calls===1?{text:'broken'}:response(b);}});const id=await upload(e);await e.start(id,{sample:true});await e.jobs.get(id).done;
 assert.equal(calls,1);assert.match((await e.status(id)).phase,/暂停/);await e.start(id,{sample:true});await e.jobs.get(id).done;assert.equal(calls,6);
});
test('pause interrupts model and restores completed extraction after restart',async()=>{
 let entered;const ready=new Promise(r=>entered=r);
 const directory=path.join(root,'pause');const e=new LargePDF({directory,classify:async(m,b,signal)=>{entered();return new Promise((resolve,reject)=>signal.addEventListener('abort',()=>reject(Error('abort')),{once:true}));}});
 const id=await upload(e);await e.start(id,{sample:true});await ready;await e.pause(id);await e.jobs.get(id).done;const s=await e.status(id);assert.equal(s.running,false);assert.equal(s.extracted,10);assert.equal(s.pending,10);
 const next=new LargePDF({directory,classify:async(m,b)=>response(b)});await next.start(id,{sample:true});await next.jobs.get(id).done;assert.equal((await next.status(id)).pending,0);
});
test('upload guards offset, size, traversal, duplicate columns and concurrent start',async()=>{
 const e=new LargePDF({directory:path.join(root,'guards'),classify:async(m,b)=>response(b)});
 await assert.rejects(()=>e.init({name:'a.pdf',size:2*1024**3+1}));await assert.rejects(()=>e.get('../a'));
 await assert.rejects(()=>e.init({name:'a.pdf',size:100,columns:'A\nA',model:{provider:'x',id:'x'}}));
 const {id}=await e.init({name:'a.pdf',size:100,model:{provider:'x',id:'x'}});await assert.rejects(()=>e.upload(id,1,Readable.from([Buffer.from('pdf')])));await assert.rejects(()=>e.start(id));
});
test('field schema rejects missing, extra and non-text fields',()=>{
 const b={columns:['金额'],records:[{id:'a'}]};for(const values of [{金额:12},{其他:'x'},{金额:'1',其他:'x'}])assert.throws(()=>validateRows(JSON.stringify([{id:'a',rows:[{values,review:''}],note:''}]),b));
});
test('binary upload accepts over 20 MB using bounded chunks and preserves exact bytes',async()=>{
 const e=new LargePDF({directory:path.join(root,'large-upload'),classify:async()=>{throw Error('not called');}});
 const chunk=Buffer.alloc(4*1024**2,93),size=24*1024**2+7;
 const {id}=await e.init({name:'large.pdf',size,model:{provider:'mock',id:'qwen3.8'}});
 for(let offset=0;offset<size;offset+=chunk.length)await e.upload(id,offset,Readable.from([chunk.subarray(0,Math.min(chunk.length,size-offset))]));
 const j=await e.get(id);assert.equal(j.uploaded,size);assert.equal((await stat(path.join(j.dir,'source.pdf'))).size,size);
 await assert.rejects(()=>e.upload(id,size,Readable.from([Buffer.from('x')])));
});
test('status can be polled while a task initializes and extracts',async()=>{
 const e=new LargePDF({directory:path.join(root,'poll'),classify:async(m,b)=>response(b)});const id=await upload(e);await e.start(id,{sample:true});
 const statuses=await Promise.all(Array.from({length:6},()=>e.status(id)));assert.ok(statuses.every(s=>s.running));await e.jobs.get(id).done;assert.equal((await e.status(id)).pending,0);
});
