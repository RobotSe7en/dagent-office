/** Independent native editable PPTX acceptance using public SDK APIs and Impress. */
import assert from "node:assert/strict";
import { execFileSync,spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { existsSync,mkdirSync,readFileSync,readdirSync,rmSync,writeFileSync } from "node:fs";
import { join,resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { createDocumentSession,openDocument,resolvePresentationGeometry } from "@dagent-office/core";
import { exportPptx,importPptx,writeDocumentZip } from "@dagent-office/office";
const root=resolve(import.meta.dirname,".."),cache=join(root,".cache/pptx-export-acceptance");
const run=(exe,args)=>execFileSync(exe,args,{cwd:root,encoding:"utf8",timeout:55_000,maxBuffer:8*1024*1024}).trim();
const digest=bytes=>createHash("sha256").update(bytes).digest("hex");
const read=path=>new Uint8Array(readFileSync(path));
const text=model=>model.snapshot().metadata.presentation.slides.map(slide=>model.getNode(slide.rootNodeId).text);
const data=model=>model.snapshot().metadata.presentation.dataSets.map(data=>({categories:data.categories,series:data.series.map(series=>({name:series.name,values:series.values}))}));
for(const name of["source","native","reopened","pdf","packages","reports"]){rmSync(join(cache,name),{recursive:true,force:true});mkdirSync(join(cache,name),{recursive:true});}
const poppler=spawnSync("pdftotext",["-v"],{encoding:"utf8",timeout:10_000});assert.equal(poppler.status,0);
const environment={node:process.version,libreOffice:run("libreoffice",["--version"]),poppler:(poppler.stdout+poppler.stderr).trim().split("\n")[0],fonts:[],powerPoint:"not-verified",wps:"not-verified",modernComments:"Exact SDK reopen and MS-PPTX structure; LibreOffice does not preserve modern threads."};
for(const family of["Noto Sans CJK SC","Liberation Sans"]){assert.ok(run("fc-match",["-f","%{family}",family]).includes(family));environment.fonts.push({family,sha256:digest(readFileSync(run("fc-match",["-f","%{file}",family])))});}
const cases=[];
for(const producer of["","producer/"])for(const name of readdirSync(join(root,"examples/pptx",producer)).filter(name=>name.endsWith(".pptx"))){
  const bytes=read(join(root,"examples/pptx",producer,name)),source=openDocument(importPptx(bytes,{documentId:"native-acceptance"})),filename=(producer?"producer-":"")+name,output=exportPptx(source),sdk=openDocument(importPptx(output.bytes,{documentId:"sdk-reopened"}));
  assert.deepEqual(text(sdk),text(source),filename+" exact SDK text");assert.deepEqual(data(sdk),data(source),filename+" exact SDK constants");assert.deepEqual(output.report.unexportedAnnotations,[]);assert.deepEqual(output.report.annotationLosses,[]);
  writeFileSync(join(cache,"source",filename),bytes);writeFileSync(join(cache,"native",filename),output.bytes);writeFileSync(join(cache,"packages",filename+".zip"),writeDocumentZip(sdk.serialize()));writeFileSync(join(cache,"reports",filename+".json"),JSON.stringify(output.report,null,2));
  cases.push({filename,source,sourceSha256:digest(bytes),exportSha256:digest(output.bytes)});
}
const sourceFiles=new Map([...readdirSync(join(root,"examples/presentation")).filter(name=>/\.(json|html|css)$/.test(name)),"assets/sample.png"].map(path=>[path,read(join(root,"examples/presentation",path))]));
const metadata=JSON.parse(new TextDecoder().decode(sourceFiles.get("metadata.json")));metadata.presentation.dataSets[0].categories=[" Q1 ","第二季度 "," 第三季度"];metadata.presentation.dataSets[0].series[0].name=" 交付数 ";sourceFiles.set("metadata.json",new TextEncoder().encode(JSON.stringify(metadata)));
const edited=createDocumentSession(openDocument(sourceFiles)),originalGeometry=resolvePresentationGeometry(edited,"heading");
edited.commit(edited.prepare({protocolVersion:1,documentId:edited.snapshot().manifest.documentId,baseRevision:0,operations:[{op:"replaceText",nodeId:"heading-p",from:0,to:4,expectedText:"项目进展",text:"项目进展（SDK 修改）"},{op:"setPresentationGeometry",nodeId:"heading",geometry:{...originalGeometry,x:60}},{op:"setStyle",nodeId:"heading-p",property:"font-size",value:"32pt"}]}));
const native=exportPptx(edited,{commentFormat:"classic"});writeFileSync(join(cache,"native/edited.pptx"),native.bytes);writeFileSync(join(cache,"reports/edited.json"),JSON.stringify(native.report,null,2));writeFileSync(join(cache,"packages/edited.zip"),writeDocumentZip(edited.serialize()));
const externalEdit=JSON.parse(run("/usr/bin/python3",[join(root,"scripts/pptx-external-edit.py"),join(cache,"native/edited.pptx"),join(cache,"reopened/external-edited.pptx")]));
const external=openDocument(importPptx(read(join(cache,"reopened/external-edited.pptx")),{documentId:"external-edit"})),p=external.snapshot().metadata.presentation;
assert.ok(text(external)[0].includes("外部编辑确认：项目进展"));assert.ok(text(external)[0].includes("外部表格"));assert.equal(p.dataSets[0].series[0].values[0],321);assert.deepEqual(p.dataSets[0].categories,metadata.presentation.dataSets[0].categories);assert.equal(p.dataSets[0].series[0].name," 交付数 ");assert.ok(p.stories.some(story=>story.kind==="notes"&&external.getNode(story.rootNodeId).text.includes("外部备注确认")));const heading=p.objects.find(object=>external.getNode(object.nodeId).text==="外部编辑确认：项目进展");assert.ok(heading);assert.ok(Math.abs(resolvePresentationGeometry(external,heading.nodeId).x-(60+200*72/2540))<.02);writeFileSync(join(cache,"reports/external-edit.json"),JSON.stringify(externalEdit,null,2));edited.dispose();
const profile=join(cache,"lo-profile-"+process.pid),lo=["-env:UserInstallation="+pathToFileURL(profile).href,"--headless"];
try{
  for(let offset=0;offset<cases.length;offset+=5){const batch=cases.slice(offset,offset+5).map(item=>join(cache,"native",item.filename));run("libreoffice",[...lo,"--convert-to","pptx","--outdir",join(cache,"reopened"),...batch]);run("libreoffice",[...lo,"--convert-to","pdf","--outdir",join(cache,"pdf"),...batch]);}
  const schemaPython=process.env.DAGENT_OFFICE_SCHEMA_PYTHON??(existsSync(join(root,".cache/pptx-schema-venv/bin/python"))?join(root,".cache/pptx-schema-venv/bin/python"):"/usr/bin/python3");
  run(schemaPython,[join(root,"scripts/check-pptx-output.py"),join(cache,"native")]);
  const records=[];
  for(const item of cases){const bytes=read(join(cache,"reopened",item.filename)),external=(()=>{try{return openDocument(importPptx(bytes,{documentId:"external-reopened"}));}catch(error){console.error(item.filename);throw error;}})(),before=item.source.snapshot(),after=external.snapshot();
    assert.equal(after.metadata.presentation.slides.length,before.metadata.presentation.slides.length,item.filename+" slides");assert.deepEqual(data(external),data(item.source),item.filename+" Office data");assert.equal(after.metadata.presentation.charts.length,before.metadata.presentation.charts.length,item.filename+" editable charts");assert.equal(after.manifest.assets.length,before.manifest.assets.length,item.filename+" images");
    const classic=before.metadata.presentation.comments.filter(comment=>comment.format==="classic").length;assert.equal(after.annotations.length,classic,item.filename+" classic comments / explicit modern limitation");
    const pdf=join(cache,"pdf",item.filename.replace(/\.pptx$/,".pdf")),pages=Number(/Pages:\s+(\d+)/.exec(run("pdfinfo",[pdf]))?.[1]);assert.equal(pages,before.metadata.presentation.slides.filter(slide=>!slide.hidden).length,item.filename+" Office PDF pages");
    records.push({filename:item.filename,sourceSha256:item.sourceSha256,exportSha256:item.exportSha256,reopenedSha256:digest(bytes),slides:after.metadata.presentation.slides.length,charts:after.metadata.presentation.charts.length,classicComments:classic,modernThreadsNotPreservedByLibreOffice:before.annotations.length-classic,pdfPages:pages});
  }
  writeFileSync(join(cache,"result.json"),JSON.stringify({environment,cases:records,externalEdit},null,2)+"\n");console.log("Native PPTX: "+records.length+" independent sources, SDK exact reopen, Office resave/PDF, and actual text/geometry/table/chart/notes edits passed.");
}finally{rmSync(profile,{recursive:true,force:true});}
