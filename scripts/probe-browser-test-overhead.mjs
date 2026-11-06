import { chromium } from '@playwright/test';
import { build } from 'vite';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import assert from 'node:assert/strict';
const directory=resolve('.cache/browser-test-overhead');mkdirSync(directory,{recursive:true});
const bundle=await build({configFile:false,build:{write:false,minify:false,lib:{entry:resolve('tests/browser/fixtures/workbook-grid-host.ts'),formats:['iife'],name:'WorkbookGridProbe'}}});
const outputs=Array.isArray(bundle)?bundle.flatMap(p=>p.output):bundle.output;
const code=outputs.filter(p=>p.type==='chunk').map(p=>p.code).join('\n');
const browser=await chromium.launch({headless:true}),cases=[];
try {
 const modes=[{host:'playground',trace:'full'},{host:'blank',trace:'full'},{host:'playground',trace:'api'},{host:'blank',trace:'api'}];
 for(let round=0;round<4;round++)for(let offset=0;offset<modes.length;offset++){
  const mode=modes[(offset+round)%modes.length];
  const context=await browser.newContext({viewport:{width:1280,height:800}});
  await context.tracing.start({snapshots:mode.trace==='full',screenshots:mode.trace==='full',sources:mode.trace==='full'});
  const page=await context.newPage();
  if(mode.host==='blank')await page.route('**/sdk-ci-overhead.html',route=>route.fulfill({contentType:'text/html',body:'<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>'}));
  const start=performance.now();
  await page.goto('http://127.0.0.1:4173/'+(mode.host==='blank'?'sdk-ci-overhead.html':''));const navigated=performance.now();
  await page.addScriptTag({content:code});await page.waitForFunction(()=>!!window.workbookGridHost);const initialized=performance.now();
  let fragments=0;
  for(let i=0;i<6;i++){
   await page.evaluate(i=>window.workbookGridHost.render({widthPt:640,heightPt:320,scrollLeftPt:20,scrollTopPt:i*20}),i);
   const result=await page.evaluate(()=>window.workbookGridHost.inspect());assert.ok(result.cellCount>0);assert.equal(result.sourceUnchanged,true);fragments+=result.cellCount;
  }
  await page.evaluate(()=>window.workbookGridHost.dispose());const actions=performance.now();
  await context.tracing.stop({path:resolve(directory,`trace-${round}-${mode.host}-${mode.trace}.zip`)});const stopped=performance.now();
  cases.push({...mode,round,navigationMs:navigated-start,initializationMs:initialized-navigated,actionsMs:actions-initialized,traceFinishMs:stopped-actions,totalMs:stopped-start,fragments});
  await context.close();
 }
 const context=await browser.newContext(),page=await context.newPage();
 await page.route('**/sdk-ci-overhead.html',route=>route.fulfill({contentType:'text/html',body:'<!doctype html><html><body></body></html>'}));await page.goto('http://127.0.0.1:4173/sdk-ci-overhead.html');
 const bytes=readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx'),array=[...bytes],base64=bytes.toString('base64'),transfer=[];
 for(let round=0;round<4;round++)for(const encoding of round%2?['base64','array']:['array','base64']){
  const start=performance.now();const length=encoding==='array'?await page.evaluate(value=>value.length,array):await page.evaluate(value=>atob(value).length,base64);assert.equal(length,bytes.length);transfer.push({round,encoding,ms:performance.now()-start,bytes:bytes.length});
 }
 await context.close();
 const report={scope:'Diagnostic only. Same six grid renders/assertions per fresh context. Local production preview, same browser; no acceptance thresholds changed.',cases,transfer};
 writeFileSync(resolve(directory,'report.json'),JSON.stringify(report,null,2)+'\n');
 const mean=(items,key)=>items.reduce((s,r)=>s+r[key],0)/items.length;
 console.log(JSON.stringify({modes:modes.map(mode=>{const items=cases.filter(r=>r.host===mode.host&&r.trace===mode.trace);return{...mode,samples:items.length,navigationMs:mean(items,'navigationMs'),initializationMs:mean(items,'initializationMs'),actionsMs:mean(items,'actionsMs'),traceFinishMs:mean(items,'traceFinishMs'),totalMs:mean(items,'totalMs')};}),transfer:['array','base64'].map(encoding=>({encoding,meanMs:mean(transfer.filter(r=>r.encoding===encoding),'ms')}))},null,2));
}finally{await browser.close();}
