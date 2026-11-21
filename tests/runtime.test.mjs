import {test} from 'node:test'
import assert from 'node:assert/strict'
import vm from 'node:vm'
import fs from 'node:fs'
import {webcrypto} from 'node:crypto'
const shared=fs.readFileSync(new URL('../src/outlook-addin/outlook-shared.js',import.meta.url),'utf8')
function setup() {
 const saved={},roaming={},handlers={},calls=[];let connectUrl=''
 const Office={
  AsyncResultStatus:{Succeeded:'ok'},
  EventType:{DialogMessageReceived:'message',DialogEventReceived:'event'},
  context:{
   roamingSettings:{get:k=>roaming[k],set:(k,v)=>roaming[k]=v,saveAsync:cb=>cb({status:'ok'})},
   ui:{displayDialogAsync:(url,opts,cb)=>{
    connectUrl=url
    cb({status:'ok',value:{addEventHandler:(k,fn)=>handlers[k]=fn,close(){}}})
   }}
  }
 }
 const ctx=vm.createContext({URL,crypto:webcrypto,console,Office,
  window:{location:{href:'https://canary.example.org:8443/outlook-addin/taskpane.html'}},
  localStorage:{getItem:k=>saved[k],setItem:(k,v)=>saved[k]=v,removeItem:k=>delete saved[k]},
  fetch:async(url,opts)=>{calls.push({url,opts});return {ok:true,json:async()=>({access_token:'synthetic-token'})}}
 })
 vm.runInContext(shared,ctx)
 return {api:ctx.canaryOutlookShared,saved,roaming,handlers,calls,getUrl:()=>new URL(connectUrl)}
}
test('OAuth dialog uses existing origin and callback; valid response persists token',async()=>{const s=setup();const p=s.api.runPluginConnect();const u=s.getUrl();assert.equal(u.origin,'https://canary.example.org:8443');assert.equal(u.pathname,'/connect/mail-plugin');assert.equal(u.searchParams.get('redirect_uri'),'https://canary.example.org:8443/outlook-addin/auth-callback.html');s.handlers.message({message:JSON.stringify({code:'synthetic-code',state:u.searchParams.get('state')})});assert.equal(await p,'synthetic-token');assert.equal(s.calls[0].url,'https://canary.example.org:8443/api/auth/plugin/token');assert.equal(s.saved.canary_outlook_addin_jwt,'synthetic-token');assert.equal(s.api.apiRoot(),'https://canary.example.org:8443/api')})
test('OAuth wrong state refuses token exchange',async()=>{const s=setup();const p=s.api.runPluginConnect();s.handlers.message({message:JSON.stringify({code:'synthetic-code',state:'wrong'})});await assert.rejects(p,/did not complete/);assert.equal(s.calls.length,0)})
test('OAuth dialog close rejects without credentials',async()=>{const s=setup();const p=s.api.runPluginConnect();s.handlers.event({error:12006});await assert.rejects(p,/closed/);assert.equal(s.calls.length,0)})
test('callback page relays code and state through Office API',()=>{const messages=[],status={};const ctx=vm.createContext({URLSearchParams,window:{location:{search:'?code=synthetic-code&state=synthetic-state'}},document:{getElementById:()=>status},Office:{onReady:fn=>fn(),context:{ui:{messageParent:value=>messages.push(JSON.parse(value))}}}});vm.runInContext(fs.readFileSync(new URL('../src/outlook-addin/auth-callback.js',import.meta.url),'utf8'),ctx);assert.deepEqual(messages,[{code:'synthetic-code',state:'synthetic-state',error:''}]);assert.match(status.textContent,/Connected/)})
