#!/usr/bin/env node
'use strict';
// OTOMASI Office agent: outbound WebSocket, Windows COM in user desktop session.
const fs=require('node:fs');
const os=require('node:os');
const path=require('node:path');
const https=require('node:https');
const http=require('node:http');
const crypto=require('node:crypto');
const {spawn}=require('node:child_process');
const WebSocket=require('ws');
const pkg=require('../package.json');
const {browserPair}=require('./browser-pair.cjs');
const ROOT=path.join(os.homedir(),'.otomasi-office');
const CREDS=path.join(ROOT,'device.json');
const BRIDGE=path.join(__dirname,'bridge.ps1');
const HUB=(process.env.OTOMASI_OFFICE_HUB||'https://mcp.otomasi.app/office').replace(/\/$/,'');
let shutting=false,connecting=false,ws=null,retry=1000;
function arg(k){let i=process.argv.indexOf(k);return i>=0?process.argv[i+1]:undefined}
function has(k){return process.argv.includes(k)}
function help(){
  console.log(`OTOMASI Office ${pkg.version}
Connect live Word, Excel and PowerPoint documents from your signed-in Windows desktop.

  npx @otomasi/office@latest                  Open default browser for pairing
  npx @otomasi/office@latest --pair <code>    Pair using one-time code
  npx @otomasi/office@latest --status         Show agent configuration
  npx @otomasi/office@latest --check          List local open documents (read-only)
  npx @otomasi/office@latest --foreground     Start outbound agent
  npx @otomasi/office@latest --logout         Delete local device credentials

First setup opens the Windows default browser to approve your Office device.
Connect https://mcp.otomasi.app/office/mcp in GPT/Claude. No manual code input required.
Requires Windows with Microsoft Office desktop and an interactive user session.
`);
}
async function request(url, method='GET', payload=null){
 return new Promise((resolve,reject)=>{
   const address=new URL(url);
   if(address.protocol!=='https:' && address.hostname!=='127.0.0.1'&&address.hostname!=='localhost')return reject(Error('HTTPS required'));
   const data=payload?Buffer.from(JSON.stringify(payload)):null;
   const mod=address.protocol==='https:'?https:http;
   const req=mod.request(address,{method,headers:{'accept':'application/json',...(data?{'Content-Type':'application/json','Content-Length':data.length}:{})},timeout:12000},res=>{
     const parts=[];res.on('data',b=>parts.push(b));res.on('end',()=>{
       let val;try{val=JSON.parse(Buffer.concat(parts).toString('utf8'));}catch{return reject(Error('Non-JSON hub response'))}
       if(res.statusCode<200||res.statusCode>=300)return reject(Error('Server HTTP '+res.statusCode+': '+(val.error||'unknown')));
       resolve(val);
     });
   });req.on('timeout',()=>req.destroy(Error('HTTP timeout')));req.on('error',reject);req.end(data);
 });
}
function readConfig(){
  try{
    const c=JSON.parse(fs.readFileSync(CREDS,'utf8'));
    if(c.hub==='https://mcp-office.otomasi.app'||c.hub==='https://office.otomasi.app'){
      c.hub=HUB;saveConfig(c);
      console.log('Migrated Office agent to:',c.hub);
    }
    return c;
  }catch{return null}
}
function saveConfig(c){fs.mkdirSync(ROOT,{recursive:true,mode:0o700});fs.writeFileSync(CREDS,JSON.stringify(c,null,2),{mode:0o600});}
function psRun(op,args={}){
 return new Promise((resolve,reject)=>{
   if(process.platform!=='win32')return reject(Error('Live Office requires Windows'));
   const data=Buffer.from(JSON.stringify({name:op,arguments:args}),'utf8').toString('base64');
   const child=spawn('powershell.exe',['-NoLogo','-NoProfile','-NonInteractive','-STA','-ExecutionPolicy','Bypass','-File',BRIDGE,'-Payload',data],
     {windowsHide:true,stdio:['ignore','pipe','pipe']});
   let out='',err='',settled=false;
   const finish=(e,value)=>{
     if(settled)return;settled=true;clearTimeout(timer);
     if(e)reject(e);else resolve(value);
   };
   const timer=setTimeout(()=>{child.kill();finish(Error('Office COM operation timeout (35s)'))},35000);
   child.stdout.on('data',b=>{out+=b.toString('utf8');if(out.length>4_000_000){child.kill();finish(Error('Output too large'))}});
   child.stderr.on('data',b=>{err+=b.toString('utf8').slice(0,500)});
   child.once('error',finish);
   child.once('close',code=>{
     let result;try{result=JSON.parse(out.trim().replace(/[\u0000-\u001f]/g,ch=>'\\u'+ch.charCodeAt(0).toString(16).padStart(4,'0')))}catch{return finish(Error('Office bridge invalid result (output_bytes='+out.length+', stderr_bytes='+err.length+', exit_code='+code+', is_json_prefix='+String(out.trim().startsWith('{'))+')'))}
     if(code!==0 || !result.ok)finish(Error(String(result.error||'Office operation failed')));
     else finish(null,result);
   });
 });
}
const names=['list','read','verify','save','word_clear','word_insert','word_replace','word_format','word_table','excel_set_range','excel_add_chart','ppt_add_slide','ppt_edit_text','word_comment','word_track_changes','excel_add_sheet','excel_delete_sheet','excel_format_range','ppt_add_textbox','ppt_delete_slide','excel_list_shapes','excel_clear_sheet','excel_freeze_rows','excel_delete_shape','ppt_add_shape','ppt_set_background','ppt_insert_image','ppt_delete_shape','excel_insert_image','word_read_paragraphs','excel_read_range','ppt_read_slide','ppt_preview_slide'];
async function pair(){
 let code=arg('--pair');
 if(!code){
   const approved=await browserPair(request,HUB,arg('--name')||os.hostname());
   saveConfig({device_id:approved.device_id,device_token:approved.device_token,hub:HUB,created_at:new Date().toISOString()});
   console.log('Office browser pairing approved.');
   return readConfig();
 }
 if(!code || !code.startsWith('office_'))throw Error('Invalid pairing code');
 const result=await request(HUB+'/device/enroll','POST',{pairing_code:code,name:arg('--name')||os.hostname(),platform:process.platform});
 saveConfig({device_id:result.device_id,device_token:result.device_token,hub:HUB,created_at:new Date().toISOString()});
 console.log('Paired successfully:',result.device_id);
 return readConfig();
}
async function serve(cfg){
 if(shutting||connecting)return;connecting=true;
 const url=(cfg.hub||HUB).replace(/^https:/,'wss:').replace(/^http:/,'ws:')+
   '/device/ws?device_id='+encodeURIComponent(cfg.device_id);
 ws=new WebSocket(url,{headers:{Authorization:'Bearer '+cfg.device_token},handshakeTimeout:12000,maxPayload:1024*1024,perMessageDeflate:false});
 let ping=null,opened=false;
 ws.on('open',()=>{
   opened=true;retry=1000;connecting=false;
   ws.send(JSON.stringify({type:'hello',tools:names.map(name=>({name})),meta:{agent_version:pkg.version,platform:process.platform}}));
   ping=setInterval(()=>{if(ws.readyState===WebSocket.OPEN)ws.send(JSON.stringify({type:'heartbeat'}))},30000);
   console.log('Office agent connected:',cfg.device_id);
 });
 ws.on('message',async raw=>{
   let msg;try{msg=JSON.parse(raw.toString('utf8'))}catch{return}
   if(msg.type==='ready')return;
   if(msg.type!=='call')return;
   const id=String(msg.id||'');
   const name=String(msg.name||'');
   if(!names.includes(name))return ws.send(JSON.stringify({type:'error',id,error:'Unsupported operation'}));
   try{
     const result=await psRun(name,msg.arguments||{});
     if(ws.readyState===WebSocket.OPEN)ws.send(JSON.stringify({type:'result',id,result}));
   }catch(error){
     if(ws.readyState===WebSocket.OPEN)ws.send(JSON.stringify({type:'error',id,error:String(error.message).slice(0,240)}));
   }
 });
 ws.on('error',err=>{console.error('Connection:',err.message)});
 ws.on('close',()=>{
   if(ping)clearInterval(ping);
   connecting=false;
   if(shutting)return;
   let pause=retry;retry=Math.min(30000,retry*2);
   console.log('Disconnected; retrying in',pause+'ms');
   setTimeout(()=>serve(cfg),pause).unref?.();
 });
}
async function main(){
 if(has('--help')||has('-h'))return help();
 if(has('--version'))return console.log(pkg.version);
 if(has('--logout')){try{fs.unlinkSync(CREDS)}catch{}return console.log('Local credentials cleared. Revoke device from MCP as well.');}
 if(has('--status')){const c=readConfig();return console.log(JSON.stringify({paired:!!c,device_id:c?.device_id||null,hub:c?.hub||HUB},null,2))}
 if(has('--check'))return console.log(JSON.stringify(await psRun('list'),null,2));
 if(process.platform!=='win32')throw Error('Current release supports Windows Office desktop. Linux server uses Python gateway.');
 let cfg=readConfig();
 if(has('--pair')||!cfg)cfg=await pair();
 process.on('SIGINT',()=>{shutting=true;ws?.close();process.exit(0)});
 process.on('SIGTERM',()=>{shutting=true;ws?.close();process.exit(0)});
 await serve(cfg);
}
main().catch(error=>{console.error('OTOMASI Office:',error.message);process.exitCode=1});
