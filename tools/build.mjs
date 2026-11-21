import fs from 'node:fs'
import path from 'node:path'
import {fileURLToPath} from 'node:url'
import {execFileSync} from 'node:child_process'
import {sha256,stageBundle,verifyOutput,normalizeOrigin} from './hosting.mjs'
const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..')
const out=process.env.CANARY_OUTLOOK_OUTPUT ? path.resolve(process.env.CANARY_OUTLOOK_OUTPUT) : path.join(root,'dist')
const names=JSON.parse(fs.readFileSync(path.join(root,'tools/runtime-files.json')))
const files={}
for (const name of names) {
  if (name.includes('..') || path.isAbsolute(name)) throw new Error('Invalid source path')
  const bytes=fs.readFileSync(path.join(root,'src',name));files[name]={sha256:sha256(bytes),content:bytes.toString('base64')}
}
const xml=fs.readFileSync(path.join(root,'src/outlook-addin/manifest.xml'),'utf8')
const version=xml.match(/<Version>([^<]+)<\/Version>/)[1]
fs.mkdirSync(out,{recursive:true})
const bundle=path.join(out,'bundle.json');fs.writeFileSync(bundle,JSON.stringify({schema:1,version,files},null,2)+'\n')
fs.copyFileSync(path.join(root,'tools/hosting.mjs'),path.join(out,'hosting.mjs'))
let commit='uncommitted',dirty=true
try { commit=execFileSync('git',['rev-parse','HEAD'],{cwd:root,encoding:'utf8',stdio:['ignore','pipe','ignore']}).trim();dirty=Boolean(execFileSync('git',['status','--porcelain'],{cwd:root,encoding:'utf8'}).trim()) } catch {}
fs.writeFileSync(path.join(out,'lock.json'),JSON.stringify({schema:1,repository:'https://github.com/canarylegal/canary-outlook',source_commit:commit,source_dirty:dirty,version,bundle_sha256:sha256(fs.readFileSync(bundle)),hosting_sha256:sha256(fs.readFileSync(path.join(out,'hosting.mjs')))},null,2)+'\n')
const origin=normalizeOrigin(process.env.CANARY_PUBLIC_URL)
stageBundle(bundle,path.join(out,'site'),origin);console.log('Verified Outlook hosting files:',verifyOutput(bundle,path.join(out,'site'),origin))
