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
    window.Office={
     AsyncResultStatus:{Succeeded:'succeeded'},
     EventType:{ItemChanged:'item',DialogMessageReceived:'message',DialogEventReceived:'event'},
     CoercionType:{Html:'html',Text:'text'},MailboxEnums:{ItemType:{Message:'message'}},HostType:{Outlook:'Outlook'},
     actions:{associate:(key,fn)=>window.testActions[key]=typeof fn},
     onReady:fn=>{if(fn)setTimeout(()=>fn({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'})},
     context:{
      requirements:{isSetSupported:()=>false},
      roamingSettings:{get:()=>null,set(){},saveAsync:cb=>cb({status:'succeeded'})},
      ui:{messageParent:value=>window.testParent.push(JSON.parse(value))},
      mailbox:{
       item:{itemType:'message',subject:'Synthetic mail',body:{getAsync:(kind,cb)=>cb({status:'succeeded',value:'Synthetic body'})}},
       userProfile:{emailAddress:'demo@example.org'},addHandlerAsync(){}
      }
     }
    }
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
  {
   const page=await browser.newPage();
   await page.addInitScript(()=>{
    localStorage.setItem('canary_outlook_addin_jwt','synthetic-token');
    window.Office={AsyncResultStatus:{Succeeded:'ok'},EventType:{ItemChanged:'item'},CoercionType:{Text:'text',Html:'html'},MailboxEnums:{ItemType:{Message:'message'}},HostType:{Outlook:'Outlook'},onReady:fn=>{if(fn)setTimeout(()=>fn({host:'Outlook'}),0);return Promise.resolve({host:'Outlook'})},actions:{associate(){}},context:{requirements:{isSetSupported:()=>false},roamingSettings:{get:()=>null,set(){},saveAsync:cb=>cb({status:'ok'})},mailbox:{item:{itemType:'message',internetMessageId:'<synthetic@example.org>',subject:'Synthetic',body:{getAsync:(k,cb)=>cb({status:'ok',value:'Body'})}},userProfile:{emailAddress:'test@example.org'},addHandlerAsync(){}}}};
   });
   await page.route('**/*',async route=>{
    const u=new URL(route.request().url());if(u.hostname==='appsforoffice.microsoft.com')return route.fulfill({contentType:'text/javascript',body:'/* synthetic Office */'});if(u.origin!=='https://canary.example.org')return route.abort();
    if(u.pathname.startsWith('/api/')) {
     let body={};
     if(u.pathname.endsWith('/cases'))body=[{id:'matter-1',case_number:'COPY-001',client_name:'Synthetic'}];
     else if(u.pathname.endsWith('/auth/me'))body={email:'test@example.org'};
     else if(u.pathname.endsWith('/linked-case')){const offset=JSON.parse(route.request().postData()).offset;body={linked_case:{id:'spoof'},requires_confirmation:true,matches:[{id:'matter-1',case_number:offset?'MORE-002':'COPY-001',client_name:'Synthetic',file_name:'<img src=x onerror="window.injected=true">.eml',folder_path:'Mail'}],next_offset:offset?null:50};}
     return route.fulfill({contentType:'application/json',body:JSON.stringify(body)});
    }
    const f=path.resolve(root,'.'+u.pathname);if(!f.startsWith(root+path.sep)||!fs.existsSync(f))return route.abort();return route.fulfill({path:f});
   });
   await page.goto('https://canary.example.org/outlook-addin/taskpane.html');
   await expect(page.locator('#filing-matches')).toContainText('COPY-001');await expect(page.locator('#btn-file')).toBeDisabled();
   await page.getByRole('button',{name:'Show more possible copies'}).click();await expect(page.locator('#filing-matches')).toContainText('MORE-002');await expect(page.locator('#btn-file')).toBeDisabled();
   assert.equal(await page.locator('#filing-matches img').count(),0);assert.equal(await page.evaluate(()=>window.injected),undefined);
   await page.getByRole('button',{name:'Choose this matter'}).first().click();await expect(page.locator('#btn-file')).toBeEnabled();
   results.push({name:'filing confirmation, pagination and safe filenames',status:'passed'});await page.close();
  }
  fs.writeFileSync(path.join(output,'results.json'),JSON.stringify(results,null,2));console.log('PASS',results.length,'synthetic Office browser checks')
 } finally {await browser.close()}
})().catch(e=>{console.error(e);process.exitCode=1})
