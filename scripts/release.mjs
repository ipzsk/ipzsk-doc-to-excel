import {readFile,mkdir,writeFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {spawnSync} from 'node:child_process';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../',import.meta.url));
const pkg=JSON.parse(await readFile(path.join(root,'package.json'),'utf8'));
if(!/^\d+\.\d+\.\d+$/.test(pkg.version))throw Error('Release requires a stable x.y.z version');
const directory=path.join(root,'release');await mkdir(directory,{recursive:true});
for(const [cmd,args] of [[process.execPath,['scripts/build.mjs']],['pnpm',['pack','--pack-destination',directory]]]){
 const r=spawnSync(cmd,args,{cwd:root,stdio:'inherit',shell:process.platform==='win32'});if(r.error||r.status!==0)throw r.error||Error('Package build failed');
}
const filename=`${pkg.name}-${pkg.version}.tgz`,buffer=await readFile(path.join(directory,filename));
await writeFile(path.join(directory,'SHA256SUMS'),`${createHash('sha256').update(buffer).digest('hex')}  ${filename}\n`);
console.log(`Release files: ${directory}`);
