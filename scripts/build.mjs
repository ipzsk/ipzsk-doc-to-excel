import {readFile,writeFile} from 'node:fs/promises';
const source=await readFile(new URL('../src/client.js',import.meta.url),'utf8');
await writeFile(new URL('../lib/client.js',import.meta.url),`window.__ModuleLoader__.load({id:"dsh-doc-to-excel",factory:(require)=>{\n${source.replace(/^export /gm,'')}\nreturn {apply,inject};\n}});\n`);
