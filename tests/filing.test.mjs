import {test} from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import vm from 'node:vm'
for (const file of ['commands.js','commands-launch.js']) {
 test(file+' never converts an identifier lookup into a send target', async()=>{
  const source=fs.readFileSync(new URL('../src/outlook-addin/'+file,import.meta.url),'utf8');const start=source.indexOf('  '+(file==='commands.js'?'async ':'')+'function resolveLinkedCase(');const end=source.indexOf('\n  function ',start+5)
  for(const body of [{linked_case:{id:'decoy'}},{matches:[{id:'one'},{id:'two'}],requires_confirmation:true}]) {
   const ctx=vm.createContext({httpFetch:async()=>({ok:true,json:async()=>body,text:JSON.stringify(body)}),apiRoot:()=>'',jsonAuthHeaders:()=>({}),parseJson:JSON.parse})
   vm.runInContext(source.slice(start,end)+'\nglobalThis.resolve=resolveLinkedCase',ctx)
   assert.equal(await ctx.resolve('token','id','mid','thread','extra'),'')
  }
 })
}
