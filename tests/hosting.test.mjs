import {test} from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import {fileURLToPath} from 'node:url'
import {execFileSync} from 'node:child_process'
import {readBundle,stageBundle,verifyOutput,normalizeOrigin,rewriteManifestXml} from '../tools/hosting.mjs'
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..')
function temp(t){const p=fs.mkdtempSync(path.join(os.tmpdir(),'outlook-test-'));t.after(()=>fs.rmSync(p,{recursive:true,force:true}));return p}
function build(output,origin=''){execFileSync(process.execPath,[path.join(root,'tools/build.mjs')],{env:{...process.env,CANARY_OUTLOOK_OUTPUT:output,CANARY_PUBLIC_URL:origin}})}
test('standalone build preserves identity, routes, query strings and every hosted dependency',t=>{
 const out=temp(t);build(out,'https://canary.example.org:8443');const bundle=readBundle(path.join(out,'bundle.json'));assert.equal(bundle.version,'1.0.28.0')
 const xml=fs.readFileSync(path.join(out,'site/outlook-addin/manifest.xml'),'utf8');assert.match(xml,/<Id>c4e8f2a1-9b3d-4a7e-8f2c-1d9e3a7b5f6a<\/Id>/)
 const active=xml.replace(/<!--[\s\S]*?-->/g,'');assert(!active.includes('YOUR_CANARY_PUBLIC_URL'))
 for(const match of active.matchAll(/https:\/\/canary\.example\.org:8443([^"<\s]*)/g)){
  const url=new URL('https://canary.example.org:8443'+match[1]);if(url.pathname!=='/')assert(fs.existsSync(path.join(out,'site',url.pathname)),url.pathname)
 }
 for(const name of Object.keys(bundle.files).filter(n=>n.endsWith('.html'))){
  const text=fs.readFileSync(path.join(out,'site',name),'utf8')
  for(const m of text.matchAll(/(?:src|href)="([^"#]+)"/g)){
   if(m[1].startsWith('https://')){assert.equal(new URL(m[1]).hostname,'appsforoffice.microsoft.com');continue}
   const target=new URL(m[1],'https://canary.example.org/'+name);assert(fs.existsSync(path.join(out,'site',target.pathname)),target.pathname)
  }
 }
 assert.equal(verifyOutput(path.join(out,'bundle.json'),path.join(out,'site'),'https://canary.example.org:8443'),Object.keys(bundle.files).length)
 const allowed=JSON.parse(fs.readFileSync(path.join(out,'site/.well-known/microsoft-officeaddins-allowed.json')));assert(allowed.allowed.every(x=>x.startsWith('https://canary.example.org:8443/outlook-addin/')))
})
test('bundle reproducibility and template staging preserve source bytes',t=>{const a=temp(t),b=temp(t);build(a);build(b);assert.deepEqual(fs.readFileSync(path.join(a,'bundle.json')),fs.readFileSync(path.join(b,'bundle.json')));const bundle=readBundle(path.join(a,'bundle.json'));for(const name of Object.keys(bundle.files))assert.deepEqual(fs.readFileSync(path.join(a,'site',name)),fs.readFileSync(path.join(root,'src',name)))})
test('renderer preserves XML comments and origin port',()=>{assert.equal(normalizeOrigin('canary.example.org:8443/'),'https://canary.example.org:8443');assert.equal(rewriteManifestXml('<!-- https://YOUR_CANARY_PUBLIC_URL -->\nhttps://YOUR_CANARY_PUBLIC_URL/outlook-addin/taskpane.html','https://canary.example.org'),'<!-- https://YOUR_CANARY_PUBLIC_URL -->\nhttps://canary.example.org/outlook-addin/taskpane.html')})
test('corrupt bundle is rejected before touching existing output',t=>{const out=temp(t);build(out);const file=path.join(out,'bundle.json');const b=JSON.parse(fs.readFileSync(file));b.files['outlook-addin/taskpane.html'].content=Buffer.from('corrupt').toString('base64');fs.writeFileSync(file,JSON.stringify(b));const before=fs.readFileSync(path.join(out,'site/outlook-addin/taskpane.html'));assert.throws(()=>stageBundle(file,path.join(out,'site')),/Invalid Outlook content/);assert.deepEqual(fs.readFileSync(path.join(out,'site/outlook-addin/taskpane.html')),before)})
test('traversal rejected; staging removes stale owned assets but preserves other CMS assets',t=>{const out=temp(t);build(out);const file=path.join(out,'bundle.json'),b=JSON.parse(fs.readFileSync(file));b.files['outlook-addin/../escape.js']=b.files['outlook-addin/taskpane.js'];const bad=path.join(out,'bad.json');fs.writeFileSync(bad,JSON.stringify(b));assert.throws(()=>stageBundle(bad,path.join(out,'site')),/Unsafe Outlook path/);fs.writeFileSync(path.join(out,'site/outlook-addin/stale.js'),'stale');fs.writeFileSync(path.join(out,'site/icons/cms-only.txt'),'keep');stageBundle(file,path.join(out,'site'));assert(!fs.existsSync(path.join(out,'site/outlook-addin/stale.js')));assert.equal(fs.readFileSync(path.join(out,'site/icons/cms-only.txt'),'utf8'),'keep')})
