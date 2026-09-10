import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp,mkdir,writeFile,readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import {newer,releaseInfo,verifyPackage,Updater} from '../lib/updater.js';
const name='dsh-doc-to-excel-0.2.2.tgz',base='https://github.com/JayceHuang/dsh-doc-to-excel/releases/download/v0.2.2/';
const release={tag_name:'v0.2.2',assets:[{name,browser_download_url:base+name},{name:'SHA256SUMS',browser_download_url:base+'SHA256SUMS'}]};
const bytes=Buffer.from('synthetic package'),sums=createHash('sha256').update(bytes).digest('hex')+'  '+name+'\n';
test('release selection validates stable version, repository asset URLs and checksums',()=>{
 assert.equal(newer('0.10.0','0.9.9'),true);assert.equal(newer('0.2.1','0.2.1'),false);
 assert.throws(()=>newer('main','0.2.1'));assert.equal(releaseInfo(release,'0.2.1').available,true);
 assert.throws(()=>releaseInfo({...release,prerelease:true},'0.2.1'));
 assert.throws(()=>releaseInfo({...release,assets:[{name,browser_download_url:'https://evil.example/file'}]},'0.2.1'));
 verifyPackage(bytes,sums,name);assert.throws(()=>verifyPackage(Buffer.from('tampered'),sums,name));
});
test('explicit install downloads, verifies, installs only this archive and verifies installed version',async()=>{
 const directory=await mkdtemp(path.join(os.tmpdir(),'dte-update-'));let calls=0;
 const u=new Updater({profileDir:directory,current:'0.2.1',fetcher:async url=>new Response(url.endsWith('/latest')?JSON.stringify(release):url.endsWith('SHA256SUMS')?sums:bytes),installer:async(profile,file)=>{calls++;assert.deepEqual(await readFile(file),bytes);await mkdir(path.join(profile,'node_modules/dsh-doc-to-excel'),{recursive:true});await writeFile(path.join(profile,'node_modules/dsh-doc-to-excel/package.json'),JSON.stringify({version:'0.2.2'}));}});
 await assert.rejects(()=>u.install('0.2.2'));await u.check();await u.install('0.2.2');await u.done;assert.equal(calls,1);assert.equal(u.state.restart,true);
});
test('checksum mismatch never invokes installer',async()=>{
 const profileDir=await mkdtemp(path.join(os.tmpdir(),'dte-update-bad-'));let calls=0;
 const u=new Updater({profileDir,current:'0.2.1',fetcher:async url=>new Response(url.endsWith('/latest')?JSON.stringify(release):url.endsWith('SHA256SUMS')?sums:'wrong'),installer:async()=>{calls++;}});
 await u.check();await u.install('0.2.2');await u.done;assert.equal(calls,0);assert.equal(u.state.error,true);
});
