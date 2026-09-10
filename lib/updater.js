import { createHash } from 'node:crypto';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { spawn } from 'node:child_process';
export const REPOSITORY='JayceHuang/dsh-doc-to-excel';
const BASE=`https://github.com/${REPOSITORY}/releases`;
export function newer(a,b){const parse=v=>{if(!/^\d+\.\d+\.\d+$/.test(v))throw Error('版本格式无效');return v.split('.').map(Number);};const av=parse(a),bv=parse(b);for(let i=0;i<3;i++){if(av[i]!==bv[i])return av[i]>bv[i];}return false;}
export function releaseInfo(data,current){
 const version=String(data.tag_name||'').replace(/^v/,'');newer(version,current);
 if(data.draft||data.prerelease)throw Error('仅支持正式发布版本');
 const filename=`dsh-doc-to-excel-${version}.tgz`,url=`${BASE}/download/v${version}/${filename}`;
 const checksumURL=`${BASE}/download/v${version}/SHA256SUMS`;
 if(!data.assets?.some(a=>a.name===filename&&a.browser_download_url===url)||!data.assets?.some(a=>a.name==='SHA256SUMS'&&a.browser_download_url===checksumURL))throw Error('此版本缺少安装包或校验文件');
 return {version,current,available:newer(version,current),filename,url,checksumURL,releaseURL:`${BASE}/tag/v${version}`};
}
export function verifyPackage(buffer,sums,filename){
 const line=sums.split('\n').find(l=>l.trim().split(/\s+\*?/)[1]===filename);
 const expected=line?.trim().split(/\s+/)[0];
 if(!/^[a-f0-9]{64}$/.test(expected||'')||createHash('sha256').update(buffer).digest('hex')!==expected)throw Error('安装包 SHA256 校验失败，未安装');
}
async function download(url,max,fetcher){
 const response=await fetcher(url,{headers:{Accept:'application/vnd.github+json','User-Agent':'dsh-doc-to-excel-updater'},signal:AbortSignal.timeout(60000)});
 if(!response.ok)throw Error(`GitHub 下载失败 (${response.status})；请检查网络或稍后重试`);
 const chunks=[];let size=0;
 for await(const chunk of response.body){size+=chunk.length;if(size>max)throw Error('更新文件超过大小限制');chunks.push(Buffer.from(chunk));}
 return Buffer.concat(chunks);
}
export function installPackage(profileDir,archive){
 const profile=path.basename(profileDir),home=path.dirname(path.dirname(profileDir));
 if(path.basename(path.dirname(profileDir))!=='profiles'||!path.isAbsolute(profileDir))throw Error('无法确定当前 profile；请使用 README 的手动更新命令');
 return new Promise((resolve,reject)=>{
   const child=spawn('dsh',['plugin','--profile',profile,'add',archive,'--ignore-scripts','--config.auto-install-peers=false'],{env:{...process.env,DSH_HOME:home},stdio:['ignore','ignore','ignore'],shell:false});
   const timer=setTimeout(()=>{child.kill();reject(Error('安装超时；请检查插件管理器状态后重试'));},180000);
   child.on('error',()=>{clearTimeout(timer);reject(Error('当前客户端找不到 dsh 命令；请使用 README 的手动更新命令'));});
   child.on('close',code=>{clearTimeout(timer);code===0?resolve():reject(Error(`插件管理器安装失败 (${code})；处理数据已保留，可重新安装之前的版本`));});
 });
}
export class Updater {
 constructor({profileDir,current,fetcher=fetch,installer=installPackage}){this.profileDir=profileDir;this.current=current;this.fetcher=fetcher;this.installer=installer;this.state={phase:'',running:false,current};}
 async check(){
   const data=JSON.parse((await download(`https://api.github.com/repos/${REPOSITORY}/releases/latest`,1024*1024,this.fetcher)).toString());
   this.release=releaseInfo(data,this.current);return {...this.release,canInstall:!!this.profileDir};
 }
 async install(version){
   if(this.state.running)throw Error('正在更新，请稍候');
   if(!this.profileDir)throw Error('未找到当前 profile，无法自动安装');
   if(!this.release||version!==this.release.version||!this.release.available)throw Error('请先检查并选择新版本');
   this.state={phase:'正在下载并校验安装包',running:true,current:this.current};
   this.done=this.perform(this.release).catch(e=>{this.state.phase=e.message;this.state.error=true;}).finally(()=>{this.state.running=false;});
   return {started:true};
 }
 async perform(release){
   const sums=await download(release.checksumURL,65536,this.fetcher);
   const buffer=await download(release.url,50*1024*1024,this.fetcher);verifyPackage(buffer,sums.toString(),release.filename);
   const directory=path.join(this.profileDir,'doc-to-excel-updates');await mkdir(directory,{recursive:true});
   const archive=path.join(directory,release.filename);await writeFile(archive,buffer);
   this.state.phase='正在安装插件，请暂时保持客户端打开';await this.installer(this.profileDir,archive);
   const installed=JSON.parse(await readFile(path.join(this.profileDir,'node_modules/dsh-doc-to-excel/package.json'),'utf8'));
   if(installed.version!==release.version)throw Error('插件仍指向开发目录，未切换到新版本；请按 README 解除开发链接后更新');
   this.state.phase=`已安装 ${release.version}，请重启客户端生效`;this.state.restart=true;
 }
}
