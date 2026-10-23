const $=id=>document.getElementById(id);
const input=$('files'),drop=$('drop'),list=$('list'),preview=$('preview'),notice=$('notice'),exportBtn=$('export');
let processing=false;
function showProgress(percent,label){const area=$('progressArea'),bar=$('progressTrack');area.hidden=false;const n=Math.max(0,Math.min(100,Math.round(percent)));$('progressFill').style.width=n+'%';bar.setAttribute('aria-valuenow',String(n));$('progressText').textContent=label+' — '+n+'%';}
function hideProgress(){ $('progressArea').hidden=true; }
const encoder=new TextEncoder(),supported=new Set(['txt','md','markdown','csv','json','pdf','docx','xlsx']),records=[];
const ext=name=>(name.split('.').pop()||'').toLowerCase();
const bytes=text=>encoder.encode(text).length;
const download=(name,blob)=>{const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000)};
function sizeLimit(){const n=Number($('limit').value);if(!Number.isInteger(n)||n<16||n>8192)throw Error('Размер части должен быть от 16 до 8192 КиБ');return n*1024}
function chunkText(text,limit){
  // Keep line boundaries when possible. Never split a Unicode code point.
  const chunks=[];let current='',currentSize=0;
  const flush=()=>{if(current){chunks.push(current);current='';currentSize=0}};
  const append=piece=>{const n=bytes(piece);if(currentSize+n>limit)flush();current+=piece;currentSize+=n};
  for(const line of text.match(/[^\n]*\n|[^\n]+$/g)||[]){
    if(bytes(line)<=limit){append(line);continue}
    flush();let piece='',size=0;
    for(const cp of line){const n=bytes(cp);if(size+n>limit){append(piece);piece='';size=0}piece+=cp;size+=n}
    if(piece)append(piece);
  }
  flush();return chunks;
}

const loadedScripts=new Map();
function loadScript(url,globalName){
  if(window[globalName])return Promise.resolve(window[globalName]);
  if(!loadedScripts.has(url))loadedScripts.set(url,new Promise((resolve,reject)=>{
    const script=document.createElement('script');script.src=url;script.onload=()=>window[globalName]?resolve(window[globalName]):reject(Error('Библиотека не инициализирована: '+globalName));
    script.onerror=()=>reject(Error('Не удалось загрузить библиотеку. Проверьте интернет.'));document.head.append(script);
  }));
  return loadedScripts.get(url);
}
async function extractDocument(file,extension,progress=()=>{}){
  if(extension==='docx'){
    const mammoth=await loadScript('https://cdn.jsdelivr.net/npm/mammoth@1.9.0/mammoth.browser.min.js','mammoth');
    const result=await mammoth.convertToHtml({arrayBuffer:await file.arrayBuffer()});
    const dom=new DOMParser().parseFromString(result.value,'text/html');
    // Preserve table cells and row boundaries instead of flattening into a single paragraph.
    for(const table of [...dom.querySelectorAll('table')]){
      const rows=[...table.querySelectorAll('tr')].map(tr=>[...tr.children].filter(c=>/^(TD|TH)$/i.test(c.tagName)).map(c=>c.textContent.trim().replace(/\s+/g,' ')).join(' | '));
      table.replaceWith(dom.createTextNode('\n[TABLE]\n'+rows.join('\n')+'\n[/TABLE]\n'));
    }
    for(const node of [...dom.querySelectorAll('h1,h2,h3,h4,h5,h6,p,li,br')]){
      if(node.tagName==='BR')node.replaceWith(dom.createTextNode('\n'));
      else node.append(dom.createTextNode('\n'));
    }
    return {text:dom.body.textContent,warning:'DOCX: изображения и сложные объединения ячеек не экспортируются; проверьте таблицы.'};
  }
  if(extension==='xlsx'){
    const XLSX=await loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','XLSX');
    const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
    const sections=book.SheetNames.map(name=>{
      const sheet=book.Sheets[name],csv=XLSX.utils.sheet_to_csv(sheet,{FS:'\t',RS:'\n',blankrows:true});
      return '===== ЛИСТ: '+name+' =====\n'+csv;
    });
    return {text:sections.join('\n\n'),warning:'XLSX: экспорт значений ячеек; формулы, форматирование, диаграммы и изображения не сохраняются.'};
  }
  if(extension==='pdf'){
    const pdfjs=await import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs');
    pdfjs.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';
    const doc=await pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;
    const pages=[];let empty=0;
    const includeImages=$('pageImages').checked;
    try{
      for(let i=1;i<=doc.numPages;i++){
        const page=await doc.getPage(i),content=await page.getTextContent();
        // PDF.js item.hasEOL and measured gaps are more reliable than inserting
        // a space between every PDF text run (which breaks words).
        const lines=[];let line='',last=null;
        for(const item of content.items){
          if(typeof item.str!=='string')continue;
          const x=item.transform?.[4]||0,y=item.transform?.[5]||0;
          if(last&&Math.abs(y-last.y)>3){lines.push(line);line='';last=null}
          if(last&&item.str){
            const gap=x-last.end;
            const fontSize=Math.max(1,Math.abs(item.transform?.[0]||item.height||10));
            if(gap>Math.max(1.5,fontSize*0.23)&&!/\s$/.test(line)&&!/^\s/.test(item.str))line+=' ';
          }
          line+=item.str;
          last={y,end:x+(item.width||0)};
          if(item.hasEOL){lines.push(line);line='';last=null}
        }
        if(line)lines.push(line);
        const text=lines.join('\n');
        if(!text.trim())empty++;
        const p={number:i,text};
        if(includeImages){
          const viewport=page.getViewport({scale:1});
          const canvas=document.createElement('canvas');canvas.width=Math.ceil(viewport.width);canvas.height=Math.ceil(viewport.height);
          const ctx=canvas.getContext('2d',{alpha:false});
          await page.render({canvasContext:ctx,viewport}).promise;
          const blob=await new Promise((resolve,reject)=>canvas.toBlob(b=>b?resolve(b):reject(Error('Не удалось создать изображение страницы '+i)),'image/jpeg',0.78));
          p.image=new Uint8Array(await blob.arrayBuffer());
          canvas.width=0;canvas.height=0;
        }
        pages.push(p);
        page.cleanup();
        progress(Math.round(i/doc.numPages*100),'PDF: обработка страницы '+i+' из '+doc.numPages);
        if(i%3===0)await new Promise(resolve=>setTimeout(resolve,0));
      }
    }finally{await doc.destroy()}
    return {text:pages.map(p=>'===== СТРАНИЦА '+p.number+' =====\n'+p.text).join('\n\n'),pages,warning:'PDF: '+pages.length+' страниц, '+empty+' без текстового слоя. '+(includeImages?'Снимки страниц включены.':'Снимки страниц отключены.')+' OCR и распознавание таблиц пока отсутствуют.'};
  }
  const text=new TextDecoder('utf-8',{fatal:true}).decode(await file.arrayBuffer());
  if(extension==='json')JSON.parse(text);
  return {text,warning:''};
}


function assessQuality(r){
  const findings=[],pages=r.pages||[];
  const add=(severity,code,message,page=null)=>findings.push({severity,code,message,page});
  if(!r.text.trim())add('error','empty_document','Не извлечён текст документа');
  if(pages.length){
    const empty=pages.filter(p=>!p.text.trim()).map(p=>p.number);
    const short=pages.filter(p=>p.text.trim()&&p.text.trim().length<40).map(p=>p.number);
    if(empty.length)add('warning','pages_without_text','Страницы без текстового слоя: '+empty.join(', '));
    if(short.length)add('info','short_pages','Страницы с малым количеством текста: '+short.join(', '));
    if(pages.some((p,i)=>p.number!==i+1))add('error','page_order','Нарушен порядок нумерации страниц');
    if(pages.some(p=>!p.image))add('info','no_page_previews','Снимки страниц отключены или отсутствуют');
    if(pages.some(p=>/\uFFFD/.test(p.text)))add('warning','replacement_characters','Обнаружены символы замены Unicode');
    const joined=pages.map(p=>'===== СТРАНИЦА '+p.number+' =====\n'+p.text).join('\n\n');
    if(joined!==r.text)add('error','page_text_mismatch','Текст страниц не совпадает с общим текстом');
  }else if(/\uFFFD/.test(r.text))add('warning','replacement_characters','Обнаружены символы замены Unicode');
  return {status:findings.some(f=>f.severity==='error')?'error':findings.some(f=>f.severity==='warning')?'review':'ok',page_count:pages.length||null,empty_page_count:pages.filter(p=>!p.text.trim()).length,preview_count:pages.filter(p=>!!p.image).length,text_bytes:bytes(r.text),findings,method:'structural-heuristics-v1',limitations:'Не проверяет точность распознавания относительно оригинала; OCR отсутствует'};
}

function render(){list.replaceChildren();exportBtn.disabled=processing||!records.some(r=>r.ok);if(!records.length){notice.textContent='Добавьте документы для обработки.';return}
notice.textContent=records.filter(r=>r.ok).length+' из '+records.length+' документов готовы.';
for(const r of records){const div=document.createElement('div');div.className='item';const info=document.createElement('div'),name=document.createElement('strong'),sub=document.createElement('small');name.textContent=r.name;sub.textContent=r.ok?'Готово · '+bytes(r.text)+' байт UTF-8'+(r.warning?' · '+r.warning:''):processing&&!r.error?'Обрабатывается…':'Ошибка: '+r.error;info.append(name,sub);div.append(info);if(r.ok){const quality=document.createElement('small');quality.textContent='Контроль качества: '+(r.quality.status==='ok'?'без замечаний':r.quality.findings.map(f=>f.message).join(' · '));info.append(quality);const b=document.createElement('button');b.textContent='Просмотр';b.onclick=()=>preview.textContent=r.text.slice(0,20000)+(r.text.length>20000?'\n…':'');div.append(b)}list.append(div)}}
async function add(files){
  if(processing)return;
  processing=true;render();
  try{
    for(let k=0;k<files.length;k++){
      const file=files[k],r={name:file.name,size:file.size,extension:ext(file.name),ok:false,error:'',text:''};
      records.push(r);render();
      const label='Файл '+(k+1)+' из '+files.length+': '+file.name;
      showProgress(0,label+' — подготовка');
      await new Promise(resolve=>setTimeout(resolve,30));
      try{
        if(!supported.has(r.extension))throw Error('формат пока не поддерживается');
        // File reading and parsing are local. For PDF, progress tracks completed pages;
        // other formats report phases rather than an invented byte-by-byte percentage.
        showProgress(5,label+' — чтение');
        const extracted=await extractDocument(file,r.extension,(p,stage)=>showProgress(5+p*0.94,label+' — '+stage));
        if(!extracted.text.trim())throw Error('не удалось извлечь текст');
        r.text=extracted.text;r.pages=extracted.pages||null;r.sourceFile=file;r.warning=extracted.warning;r.quality=assessQuality(r);r.ok=r.quality.status!=='error';if(!r.ok)throw Error('Проверка качества: '+r.quality.findings.filter(f=>f.severity==='error').map(f=>f.message).join('; '));
        showProgress(100,label+' — готово');
      }catch(e){r.error=e.message||String(e);showProgress(100,label+' — ошибка')}
      render();
      await new Promise(resolve=>setTimeout(resolve,0));
    }
  }finally{processing=false;render();/* Keep the final 100% or error status visible until next upload or clear. */}
}

const crcTable=Uint32Array.from({length:256},(_,i)=>{let c=i;for(let k=0;k<8;k++)c=(c>>>1)^((c&1)?0xedb88320:0);return c>>>0});
function crc32(data){let crc=0xffffffff;for(let i=0;i<data.length;i++)crc=crcTable[(crc^data[i])&255]^(crc>>>8);return (crc^0xffffffff)>>>0}
const u16=(view,pos,value)=>view.setUint16(pos,value,true),u32=(view,pos,value)=>view.setUint32(pos,value>>>0,true);
async function zipStore(files){
  // Compress only textual metadata losslessly; keep JPEG and PDF bytes untouched.
  const parts=[],central=[];let offset=0;
  for(const f of files){const name=encoder.encode(f.name),data=f.data instanceof Uint8Array?f.data:(f.content instanceof Uint8Array?f.content:encoder.encode(f.content)),crc=crc32(data);
    let packed=data,method=0;
    if(f.name.endsWith('.json')&&data.length>256&&typeof CompressionStream!=='undefined'){
      try{const stream=new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));const result=new Uint8Array(await new Response(stream).arrayBuffer());if(result.length<data.length){packed=result;method=8}}catch(_){/* Browser fallback: ZIP STORE */}
    }
    if(name.length>65535||data.length>0xffffffff)throw Error('Слишком большой элемент ZIP');
    const lh=new Uint8Array(30+name.length),lv=new DataView(lh.buffer);
    u32(lv,0,0x04034b50);u16(lv,4,20);u16(lv,6,0x0800);u16(lv,8,method);u32(lv,14,crc);u32(lv,18,packed.length);u32(lv,22,data.length);u16(lv,26,name.length);lh.set(name,30);
    parts.push(lh,packed);
    const ch=new Uint8Array(46+name.length),cv=new DataView(ch.buffer);
    u32(cv,0,0x02014b50);u16(cv,4,20);u16(cv,6,20);u16(cv,8,0x0800);u16(cv,10,method);u32(cv,16,crc);u32(cv,20,packed.length);u32(cv,24,data.length);u16(cv,28,name.length);u32(cv,42,offset);ch.set(name,46);central.push(ch);
    offset+=lh.length+packed.length;
  }
  if(files.length>65535)throw Error('Слишком много частей для ZIP');
  const centralSize=central.reduce((n,p)=>n+p.length,0);
  if(offset+centralSize>0xffffffff)throw Error('ZIP больше 4 ГиБ: необходим ZIP64');
  const end=new Uint8Array(22),v=new DataView(end.buffer);u32(v,0,0x06054b50);u16(v,8,files.length);u16(v,10,files.length);u32(v,12,centralSize);u32(v,16,offset);
  return new Blob([...parts,...central,end],{type:'application/zip'});
}

function findHeadings(text){
  // Conservative detection: exclude numbered instructions, page headers,
  // dotted table-of-contents entries and long sentences.
  const excluded=/^(?:step(?:s)?|note|figure|table|contents|before you start|hikrobot idmvs client software user manual)$/i;
  return text.split('\n').map(x=>x.trim()).filter(x=>{
    if(x.length<5||x.length>100||excluded.test(x))return false;
    if(/\.{3,}\s*\d*\s*$/.test(x)||/[.!?:;]\s*$/.test(x))return false;
    if(/^[1-9]\d*\.\s+/.test(x))return false; // numbered procedure steps
    if(/^(?:Chapter|Section)\s+\d+\b/i.test(x))return true;
    if(/^\d+(?:\.\d+)+\s+\S/.test(x))return true; // 6.6.2 Algorithm ROI
    return /^(?:АННОТАЦИЯ|ВВЕДЕНИЕ|ОГЛАВЛЕНИЕ|ПОДГОТОВКА|ВОЗМОЖНЫЕ ОШИБКИ)$/iu.test(x);
  }).slice(0,25);
}

function makePdfChunks(record,limit){
  const units=record.pages.map((p,i)=>({
    page:p.number,
    text:(i?'\n\n':'')+'===== СТРАНИЦА '+p.number+' =====\n'+p.text,
    headings:findHeadings(p.text)
  }));
  const result=[];
  for(const unit of units){
    const segments=chunkText(unit.text,Math.max(1024,limit-8192));
    for(let i=0;i<segments.length;i++)result.push({
      text:segments[i],page_start:unit.page,page_end:unit.page,
      headings:unit.headings,segment_index:i+1,segment_count:segments.length
    });
  }
  if(result.map(c=>c.text).join('')!==record.text)throw Error('Контроль PDF: текст после разделения отличается от исходного извлечения');
  return result;
}

const hex=buf=>[...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,'0')).join('');
function safeDocumentStem(name){
  const map={а:'a',б:'b',в:'v',г:'g',д:'d',е:'e',ё:'e',ж:'zh',з:'z',и:'i',й:'y',к:'k',л:'l',м:'m',н:'n',о:'o',п:'p',р:'r',с:'s',т:'t',у:'u',ф:'f',х:'kh',ц:'ts',ч:'ch',ш:'sh',щ:'sch',ъ:'',ы:'y',ь:'',э:'e',ю:'yu',я:'ya'};
  const stem=name.replace(/\.[^.]+$/,'').normalize('NFKD').toLowerCase();
  return (Array.from(stem,c=>map[c]??c).join('').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_-]+/g,'_').replace(/_+/g,'_').replace(/^[_-]+|[_-]+$/g,'').slice(0,90)||'document');
}
async function makePackage(){
  const limit=sizeLimit(),ready=records.filter(r=>r.ok);if(!ready.length)throw Error('Нет готовых документов');
  const files=[],documents=[],qualityReports=[];
  for(let i=0;i<ready.length;i++){
    const r=ready[i],id='doc-'+String(i+1).padStart(4,'0'),stem=safeDocumentStem(r.name),folder=id+'_'+stem,chunks=r.extension==='pdf'?makePdfChunks(r,limit):chunkText(r.text,Math.max(1024,limit-4096)).map(text=>({text,page_start:null,page_end:null,headings:[]}));
    const sha=hex(await crypto.subtle.digest('SHA-256',encoder.encode(r.text)));
    qualityReports.push({document_id:id,source_name:r.name,...r.quality});
    const parts=[];
    for(let j=0;j<chunks.length;j++){
      const path=folder+'/'+stem+'_part_'+String(j+1).padStart(4,'0')+'.json';
      const payload={schema:'okto-chunk-v1',document_id:id,source_name:r.name,part_index:j+1,part_count:chunks.length,text:chunks[j].text,page_start:chunks[j].page_start,page_end:chunks[j].page_end,headings:chunks[j].headings,image_paths:chunks[j].page_start&&r.pages.find(p=>p.number===chunks[j].page_start)?.image?[folder+'/pages/page-'+String(chunks[j].page_start).padStart(4,'0')+'.jpg']:[]};
      const content=JSON.stringify(payload);if(bytes(content)>limit)throw Error('Часть превысила лимит: '+path);
      files.push({name:path,content});parts.push({path,index:j+1,text_bytes:bytes(chunks[j].text),page_start:chunks[j].page_start,page_end:chunks[j].page_end,headings:chunks[j].headings});
    }
    let sourcePath=null;
    if(r.extension==='pdf'&&$('keepSource').checked){
      sourcePath=folder+'/source.pdf';
      files.push({name:sourcePath,content:new Uint8Array(await r.sourceFile.arrayBuffer())});
    }
    const pageMap=r.pages?r.pages.map(p=>{
      const pageNumber=p.number??p.page;
      const imagePath=p.image?folder+'/pages/page-'+String(pageNumber).padStart(4,'0')+'.jpg':null;
      if(p.image)files.push({name:imagePath,data:p.image});
      return {page:pageNumber,marker:'===== СТРАНИЦА '+pageNumber+' =====',text_bytes:bytes(p.text),image_path:imagePath,has_text:!!p.text.trim()};
    }):null;
    if(r.extension==='pdf'&&$('pageImages').checked){
      const expected=r.pages?.length||0;
      const actual=pageMap?.filter(p=>p.image_path).length||0;
      if(!expected||actual!==expected)throw Error('Контроль PDF: снимки страниц отсутствуют ('+actual+' из '+expected+'). Перезагрузите страницу и повторите обработку.');
    }
    if(pageMap?.some(p=>!Number.isInteger(p.page)))throw Error('Контроль PDF: неверная нумерация страниц');
    documents.push({document_id:id,source_name:r.name,source_type:r.extension,extraction_warning:r.warning||null,source_bytes:r.size,source_path:sourcePath,page_map:pageMap,text_bytes:bytes(r.text),text_sha256:sha,part_count:parts.length,parts});
  }
  const manifest={schema:'okto-package-v1',created_at:new Date().toISOString(),encoding:'UTF-8',split_strategy:'pdf-page-boundaries-with-line-fallback; other-formats-line-preferred',pdf_pages:'page markers, chunk page metadata, image_paths, page_map',documents};
  for(const doc of documents){
    for(const p of doc.page_map||[]){
      if(p.image_path&&!files.some(f=>f.name===p.image_path))throw Error('Изображение отсутствует в ZIP: '+p.image_path);
    }
  }
  manifest.quality_report_path='quality-report.json';
  files.unshift({name:'quality-report.json',content:JSON.stringify({schema:'okto-quality-v1',documents:qualityReports},null,2)});
  files.unshift({name:'manifest.json',content:JSON.stringify(manifest,null,2)});
  return {blob:await zipStore(files),parts:files.length-1,archiveName:ready.length===1?safeDocumentStem(ready[0].name)+'_split.zip':'OKTO_documents_split.zip'};
}
drop.onclick=()=>input.click();drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click()}};
input.onchange=()=>{add([...input.files]);input.value=''};
drop.ondragover=e=>{e.preventDefault();drop.classList.add('over')};drop.ondragleave=()=>drop.classList.remove('over');
drop.ondrop=e=>{e.preventDefault();drop.classList.remove('over');add([...e.dataTransfer.files])};
$('clear').onclick=()=>{if(processing)return;records.length=0;hideProgress();preview.textContent='Выберите документ.';render()};
exportBtn.onclick=async()=>{exportBtn.disabled=true;notice.textContent='Формируем ZIP…';showProgress(0,'Упаковка файлов в ZIP');await new Promise(resolve=>setTimeout(resolve,30));try{const {blob,parts,archiveName}=await makePackage();showProgress(100,'ZIP сформирован');download(archiveName,blob);notice.textContent='ZIP сформирован: '+parts+' частей. Импорт в бот пока не реализован.'}catch(e){notice.textContent='Ошибка экспорта: '+(e.message||String(e));showProgress(100,'Ошибка упаковки ZIP')}finally{exportBtn.disabled=false}};
render();