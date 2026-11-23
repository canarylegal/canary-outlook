import {test} from 'node:test'
import assert from 'node:assert/strict'
import fs from 'node:fs'
import vm from 'node:vm'
const source=fs.readFileSync(new URL('../src/outlook-addin/taskpane.js',import.meta.url),'utf8')
const body=source.slice(source.indexOf('  async function fileMessageToCase('),source.indexOf('  async function refreshAuthAndCases('))
for (const pendingPart of ['parent','attachment']) test('pending '+pendingPart+' uploads never get a filed category',async()=>{
 const shown=[];let calls=0
 const item={itemType:'message',attachments:[{id:'a',name:'a.txt',contentType:'text/plain'}]}
 const context=vm.createContext({Office:{context:{mailbox:{item}},MailboxEnums:{ItemType:{Message:'message'}},CoercionType:{Text:'text'}},
  show:(...args)=>shown.push(args),getBodyAsync:async()=> 'body',$:()=>null,mailItemSubject:async()=> 'subject',sanitizeFilename:x=>x,
  buildSyntheticEml:()=>new Blob(['body']),primaryOutlookItemIdForApi:()=>'',shouldSkipAttachment:()=>false,getAttachmentContentAsync:async()=>({content:'YQ=='}),base64ToBlob:()=>new Blob(['a']),
  uploadMultipart:async()=>({id:'file-'+(++calls),scan_state:(pendingPart==='parent'?calls===1:calls===2)?'pending':'clean'}),
  applyCanaryOutlookCategoryWithRetries:()=>{throw Error('Must not tag pending upload')},refreshLinkedCaseForCurrentItem:()=>{},Blob})
 vm.runInContext(body+'\nglobalThis.file=fileMessageToCase',context)
 await context.file('token','case');assert.equal(calls,2)
 assert.ok(shown.some(x=>String(x[1]).includes('security check pending')))
})
