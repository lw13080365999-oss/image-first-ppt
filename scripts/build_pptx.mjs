#!/usr/bin/env node
// Executes a GPT-authored object map. This file does not design slide layouts.
import fs from 'node:fs/promises';
import path from 'node:path';
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
import crypto from 'node:crypto';

const req = createRequire(import.meta.url);
function fail(message) { throw new Error(message); }
function flag(name) { const i=process.argv.indexOf(name); return i<0?undefined:process.argv[i+1]; }
function local(root, rel) {
  if(typeof rel!=='string'||!rel||path.isAbsolute(rel)) fail('Expected project-relative path: '+rel);
  const p=path.resolve(root,rel);
  if(!p.startsWith(root+path.sep)) fail('Path escapes project: '+rel);
  return p;
}
async function safeRead(root, rel) {
  const p=local(root,rel), real=await fs.realpath(p);
  if(!real.startsWith(root+path.sep)) fail('Symlink escapes project: '+rel);
  return fs.readFile(real);
}
const readJSON=async p=>JSON.parse(await fs.readFile(p,'utf8'));
const digest=b=>crypto.createHash('sha256').update(b).digest('hex');
function pos(bbox) {
  if(!Array.isArray(bbox)||bbox.length!==4||!bbox.every(Number.isFinite)||bbox[2]<=0||bbox[3]<=0) fail('bbox must be [x,y,w,h] in inches with positive size');
  return {left:bbox[0]*96,top:bbox[1]*96,width:bbox[2]*96,height:bbox[3]*96};
}
function content(slide,ref) {
  if(ref==='title') return {text:slide.title};
  const [kind,id]=String(ref??'').split(':');
  const field={text:'texts',table:'tables',chart:'charts',formula:'formulas'}[kind];
  const item=field&&(slide[field]??[]).find(x=>x.id===id);
  if(!item) fail('Unknown content_ref '+ref+' on '+slide.id);
  return item;
}
function textStyle(o,font) {
  return {typeface:o.font_family??font,fontSize:(o.font_size_pt??22)*96/72,
    color:o.color??'#172536',bold:o.bold??false,italic:o.italic??false,
    alignment:o.align??'left',verticalAlignment:o.valign??'top',wrap:'square',
    autoFit:'none',insets:0};
}
async function main() {
  if(process.argv.includes('--help')) {
    console.log('Usage: node build_pptx.mjs --project DIR --output draft.pptx [--check-runtime]\nPPT_NODE_MODULES may point to the installed @oai/artifact-tool node_modules directory. Output is project-relative and must not exist.');return;
  }
  const modulePath=req.resolve('@oai/artifact-tool',{paths:process.env.PPT_NODE_MODULES?[process.env.PPT_NODE_MODULES]:undefined});
  if(process.argv.includes('--check-runtime')) { console.log(JSON.stringify({artifact_tool:modulePath,node:process.version}));return; }
  const {Presentation,PresentationFile}=await import(pathToFileURL(modulePath).href);
  const root=await fs.realpath(path.resolve(flag('--project')??fail('--project required')));
  const deck=await readJSON(path.join(root,'deck.json'));
  if(deck.schema_version!==1||!Array.isArray(deck.slides)||!deck.slides.length) fail('Need non-empty deck.json schema_version=1');
  const output=local(root,flag('--output')??'build/draft.pptx');
  try { await fs.access(output);fail('Refusing to overwrite '+output); } catch(e) {if(e.code!=='ENOENT') throw e;}
  await fs.mkdir(path.dirname(output),{recursive:true});
  const outputParent=await fs.realpath(path.dirname(output));
  if(outputParent!==root&&!outputParent.startsWith(root+path.sep)) fail('Output directory escapes project');
  const presentation=Presentation.create({slideSize:{width:deck.canvas.width_in*96,height:deck.canvas.height_in*96}});
  const font=deck.style?.font_family??'Arial';
  presentation.theme.defaultFont=font;
  const assets=new Map((deck.assets??[]).map(a=>[a.id,a]));
  const report={schema_version:1,engine:'@oai/artifact-tool',slides:[],native_formulas_pending:[],design_stage:'external image model; object-map executor does not generate designs'};
  for(const spec of deck.slides) {
    if(!/^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/.test(spec.id)) fail('Unsafe slide id');
    const map=JSON.parse(await safeRead(root,`slides/${spec.id}/objects.json`));
    if(map.slide_id!==spec.id||!map.design?.path||!map.design?.sha256) fail('Object map requires matching slide_id and design path/hash');
    const design=await safeRead(root,map.design.path);
    if(digest(design)!==map.design.sha256) fail('Design changed after mapping: '+spec.id);
    const slide=presentation.slides.add();
    slide.background.fill=map.background??deck.style?.background??'#FFFFFF';
    const built=new Map(), ids=new Set(), row={id:spec.id,objects:[]};
    for(const o of map.objects??[]) {
      if(!o.id||ids.has(o.id)) fail('Missing/duplicate object id on '+spec.id);ids.add(o.id);
      if(o.kind==='connector') continue;
      const position=pos(o.bbox), name=`ifppt:${spec.id}:${o.id}`;
      let object;
      if(o.kind==='text'||o.kind==='shape'||o.kind==='formula') {
        object=slide.shapes.add({geometry:o.kind==='shape'?(o.geometry??'rect'):'textbox',name,position,
          fill:o.fill??'none',line:{fill:o.line_color??'none',width:(o.line_width_pt??0)*96/72,style:'solid'}});
        if(o.kind==='formula') {
          const f=content(spec,o.content_ref);
          if(!o.omml_path) fail('Native formula requires omml_path: '+o.id);
          await safeRead(root,o.omml_path);
          object.text=`IFPPT_FORMULA_${f.id}`;
          report.native_formulas_pending.push({slide_id:spec.id,object_id:o.id,omml_path:o.omml_path});
        } else {
          const v=o.content_ref?content(spec,o.content_ref).text:(o.text??'');
          if(o.kind==='text'&&!o.content_ref) fail('Text object must bind source content: '+o.id);
          object.text=String(v);
        }
        object.text.style=textStyle(o,font);
      } else if(o.kind==='image') {
        const a=o.content_ref?.startsWith('asset:')?assets.get(o.content_ref.slice(6)):null;
        if(o.content_ref?.startsWith('asset:')&&!a) fail('Unknown image asset: '+o.content_ref);
        if(a&&o.path&&o.path!==a.path) fail('Image path conflicts with source asset: '+o.id);
        const rel=a?.path??o.path;if(!rel) fail('Image path missing: '+o.id);
        const bytes=await safeRead(root,rel);
        const ext=path.extname(rel).toLowerCase();
        const type={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg','.svg':'image/svg+xml','.webp':'image/webp'}[ext];
        if(!type) fail('Unsupported image type '+ext);
        object=slide.images.add({blob:bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),contentType:type,
          position,fit:o.fit??'contain',alt:o.alt??a?.id??o.id});
      } else if(o.kind==='table') {
        const t=content(spec,o.content_ref),values=t.cells;
        if(!Array.isArray(values)||!values.length||!values[0].length||values.some(r=>r.length!==values[0].length)) fail('Table must be rectangular');
        object=slide.tables.add({rows:values.length,columns:values[0].length,...position,values:values.map(r=>r.map(String))});
        object.borders.assign({fill:o.line_color??'#D4DCE4',width:1,style:'solid'});
        for(let r=0;r<values.length;r++) for(let c=0;c<values[0].length;c++) {
          const cell=object.getCell(r,c);cell.text.style=textStyle({...o,bold:r===0?true:o.bold,font_size_pt:o.font_size_pt??17},font);
          cell.fill=r===0?(o.header_fill??'#E8EEF3'):(o.fill??'#FFFFFF');
        }
      } else if(o.kind==='chart') {
        const c=content(spec,o.content_ref);
        object=slide.charts.add(o.chart_type??'bar',{
          hasLegend:(c.series??[]).length>1,dataLabels:{showValue:true},...(o.chart_options??{}),
          position,categories:c.categories,series:c.series});
      } else fail('Unsupported object kind '+o.kind);
      built.set(o.id,object);row.objects.push({id:o.id,kind:o.kind,content_ref:o.content_ref??null,name});
    }
    for(const o of map.objects??[]) if(o.kind==='connector') {
      if(!built.has(o.from)||!built.has(o.to)) fail('Connector endpoint missing: '+o.id);
      slide.shapes.connect(built.get(o.from),built.get(o.to),{kind:o.connector_kind??'straight',fromSide:o.from_side??'right',toSide:o.to_side??'left',
        line:{fill:o.line_color??'#334155',width:(o.line_width_pt??1.5)*96/72},...(o.arrow===false?{}:{tail:{type:'arrow'}})});
      row.objects.push({id:o.id,kind:'connector',from:o.from,to:o.to});
    }
    slide.speakerNotes.text=spec.notes??'';
    report.slides.push(row);
  }
  await (await PresentationFile.exportPptx(presentation)).save(output);
  const bytes=await fs.readFile(output);report.output=path.relative(root,output);report.sha256=digest(bytes);
  const reportPath=output+'.build.json';await fs.writeFile(reportPath,JSON.stringify(report,null,2));
  console.log(JSON.stringify({draft:output,report:reportPath,native_formulas_pending:report.native_formulas_pending.length,final_review_required:true}));
}
main().catch(e=>{console.error(e.message??e);process.exitCode=1;});
