'use strict';
const os=require('node:os');
const {spawn}=require('node:child_process');
const sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
function openDefaultBrowser(link){
 try{
  // Windows delegates URLs to the user's registered default browser.
  const child=spawn('rundll32.exe',['url.dll,FileProtocolHandler',link],
    {detached:true,stdio:'ignore',windowsHide:true});
  child.once('error',()=>{console.log('Please open this link in your browser:',link)});
  child.unref();
  return true;
 }catch{return false}
}
async function browserPair(request,hub,name){
 const started=await request(hub+'/device/browser/start','POST',{
   name:name||os.hostname(),platform:'win32'});
 const code=started.user_code, link=started.verification_uri_complete;
 if(!code||!link)throw Error('Hub returned incomplete browser pairing details');
 console.log('\nOTOMASI OFFICE · BROWSER APPROVAL');
 console.log('Computer: '+(name||os.hostname()));
 console.log('Your verification code: '+code);
 console.log('Open to approve: '+link);
 console.log('Only approve if the computer and code match. Expires in '+started.expires_in+' seconds.');
 if(!openDefaultBrowser(link))console.log('Automatic browser launch unavailable. Copy the link above.');
 const stop=Date.now()+(started.expires_in||600)*1000;
 let lastError='';
 while(Date.now()<stop){
   await sleep(Math.max(2000,(started.interval||3)*1000));
   try{
     const result=await request(hub+'/device/browser/token','POST',{device_code:started.device_code});
     if(result.status==='authorization_pending')continue;
     if(result.device_id&&result.device_token)return result;
     if(result.error)throw Error(result.error);
     lastError='Unexpected pairing server reply';
   }catch(err){
     if(/expired|invalid_device_code|account_disabled|office_device_limit/.test(err.message))throw err;
     lastError=err.message;
   }
 }
 throw Error('Browser authorization expired. '+(lastError?('Last issue: '+lastError):'Run npx again.'));
}
module.exports={browserPair};
