// Optional synthetic Office browser checks. Use NODE_PATH for an external Playwright installation.
const {chromium,expect}=require('@playwright/test')
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict')
const root=path.resolve(__dirname,'../dist/site'),output=process.env.CANARY_OUTLOOK_BROWSER_OUTPUT||path.resolve(__dirname,'../dist/browser')
fs.mkdirSync(output,{recursive:true})
;(async()=>{
 const browser=await chromium.launch({headless:true}),results=[]
 try {
  for(const name of ['taskpane','compose-pane','commands','auth-callback']){
   const page=await browser.newPage(),errors=[],requests=[]
   page.on('pageerror',e=>errors.push(e.message))
   await page.addInitScript(()=>{
    window.testParent=[];window.testActions={}
    window.Office={AsyncResultStatus:{Succeeded:'succeeded'},EventType:{ItemChanged:'item',DialogMessageReceived:'message',DialogEventReceived:'event'},CoercionType:{Html:'html',Text:'text'},MailboxEnums:{ItemType:{Message:'message'}},HostType:{Outlook:'Outlook'},actions:{associate:(key,fn)=>window.testActions[key]=typeof fn},onReady:fn=>{if(fn)setTimeout(()=>fn({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'})},context:{requirements:{isSetSupported:()=>false},roamingSettings:{get:()=>null,set(){},saveAsync:cb=>cb({status:'succeeded'})},ui:{messageParent:value=>window.testParent.push(JSON.parse(value))},mailbox:{item:{itemType:'message',subject:'Synthetic mail',body:{getAsync:(kind,cb)=>cb({status:'succeeded',value:'Synthetic body'})},userProfile:{emailAddress:'demo@example.org'},addHandlerAsync(){}}}}
   })
   await page.route('**/*',async route=>{
    const url=new URL(route.request().url());requests.push(url.pathname)
    if(url.hostname==='appsforoffice.microsoft.com')return route.fulfill({contentType:'text/javascript',body:'/* Synthetic Office API supplied by fixture. */'})
    if(url.origin!=='https://canary.example.org')return route.abort()
    const file=path.resolve(root,'.'+url.pathname)
    if(!file.startsWith(root+path.sep)||!fs.existsSync(file))return route.abort()
    return route.fulfill({path:file})
   })
   await page.goto('https://canary.example.org/outlook-addin/'+name+'.html?code=synthetic-code&state=synthetic-state')
   if(name==='taskpane'||name==='compose-pane')await expect(page.locator('#btn-sign-in')).toBeVisible()
   if(name==='auth-callback'){
    await expect(page.locator('#status')).toContainText('Connected')
    assert.deepEqual(await page.evaluate(()=>window.testParent),[{code:'synthetic-code',state:'synthetic-state',error:''}])
   }
   if(name==='commands')assert(Object.keys(await page.evaluate(()=>window.testActions)).length>0)
   assert.deepEqual(errors,[],name+' page errors')
   assert(!requests.some(x=>x.startsWith('/api/')),'Signed-out page should not contact APIs')
   results.push({name,status:'passed',requests});await page.close()
  }
  fs.writeFileSync(path.join(output,'results.json'),JSON.stringify(results,null,2));console.log('PASS',results.length,'synthetic Office browser checks')
 } finally {await browser.close()}
})().catch(e=>{console.error(e);process.exitCode=1})
