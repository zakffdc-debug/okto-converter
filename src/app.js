const $=id=>document.getElementById(id);
const input=$('files'),drop=$('drop'),list=$('list'),preview=$('preview'),notice=$('notice'),exportBtn=$('export');
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
async function extractDocument(file,extension){
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
      }
    }finally{await doc.destroy()}
    return {text:pages.map(p=>'===== СТРАНИЦА '+p.number+' =====\n'+p.text).join('\n\n'),pages,warning:'PDF: '+pages.length+' страниц, '+empty+' без текстового слоя. '+(includeImages?'Снимки страниц включены.':'Снимки страниц отключены.')+' OCR и распознавание таблиц пока отсутствуют.'};
  }
  const text=new TextDecoder('utf-8',{fatal:true}).decode(await file.arrayBuffer());
  if(extension==='json')JSON.parse(text);
  return {text,warning:''};
}

function render(){list.replaceChildren();exportBtn.disabled=!records.some(r=>r.ok);if(!records.length){notice.textContent='Добавьте документы для обработки.';return}
notice.textContent=records.filter(r=>r.ok).length+' из '+records.length+' документов готовы.';
for(const r of records){const div=document.createElement('div');div.className='item';const info=document.createElement('div'),name=document.createElement('strong'),sub=document.createElement('small');name.textContent=r.name;sub.textContent=r.ok?'Готово · '+bytes(r.text)+' байт UTF-8'+(r.warning?' · '+r.warning:''):'Ошибка: '+r.error;info.append(name,sub);div.append(info);if(r.ok){const b=document.createElement('button');b.textContent='Просмотр';b.onclick=()=>preview.textContent=r.text.slice(0,20000)+(r.text.length>20000?'\n…':'');div.append(b)}list.append(div)}}
async function add(files){for(const file of files){const r={name:file.name,size:file.size,extension:ext(file.name),ok:false,error:'',text:''};records.push(r);render();try{if(!supported.has(r.extension))throw Error('формат пока не поддерживается');const extracted=await extractDocument(file,r.extension);if(!extracted.text.trim())throw Error('не удалось извлечь текст');r.text=extracted.text;r.pages=extracted.pages||null;r.sourceFile=file;r.warning=extracted.warning;r.ok=true}catch(e){r.error=e.message||String(e)}render()}}
function crc32(data){let crc=-1;for(const b of data){crc^=b;for(let k=0;k<8;k++)crc=(crc>>>1)^((crc&1)?0xedb88320:0)}return (crc^-1)>>>0}
const u16=(view,pos,value)=>view.setUint16(pos,value,true),u32=(view,pos,value)=>view.setUint32(pos,value>>>0,true);
function zipStore(files){
  // ZIP STORE, no compression; avoids external libraries and network dependencies.
  const parts=[],central=[];let offset=0;
  for(const f of files){const name=encoder.encode(f.name),data=f.content instanceof Uint8Array?f.content:encoder.encode(f.content),crc=crc32(data);
    if(name.length>65535||data.length>0xffffffff)throw Error('Слишком большой элемент ZIP');
    const lh=new Uint8Array(30+name.length),lv=new DataView(lh.buffer);
    u32(lv,0,0x04034b50);u16(lv,4,20);u16(lv,6,0x0800);u16(lv,8,0);u32(lv,14,crc);u32(lv,18,data.length);u32(lv,22,data.length);u16(lv,26,name.length);lh.set(name,30);
    parts.push(lh,data);
    const ch=new Uint8Array(46+name.length),cv=new DataView(ch.buffer);
    u32(cv,0,0x02014b50);u16(cv,4,20);u16(cv,6,20);u16(cv,8,0x0800);u32(cv,16,crc);u32(cv,20,data.length);u32(cv,24,data.length);u16(cv,28,name.length);u32(cv,42,offset);ch.set(name,46);central.push(ch);
    offset+=lh.length+data.length;
  }
  if(files.length>65535)throw Error('Слишком много частей для ZIP');
  const centralSize=central.reduce((n,p)=>n+p.length,0);
  if(offset+centralSize>0xffffffff)throw Error('ZIP больше 4 ГиБ: необходим ZIP64');
  const end=new Uint8Array(22),v=new DataView(end.buffer);u32(v,0,0x06054b50);u16(v,8,files.length);u16(v,10,files.length);u32(v,12,centralSize);u32(v,16,offset);
  return new Blob([...parts,...central,end],{type:'application/zip'});
}
const hex=buf=>[...new Uint8Array(buf)].map(b=>b.toString(16).padStart(2,'0')).join('');
async function makePackage(){
  const limit=sizeLimit(),ready=records.filter(r=>r.ok);if(!ready.length)throw Error('Нет готовых документов');
  const files=[],documents=[];
  for(let i=0;i<ready.length;i++){
    const r=ready[i],id='doc-'+String(i+1).padStart(4,'0'),chunks=chunkText(r.text,Math.max(1024,limit-4096));
    const sha=hex(await crypto.subtle.digest('SHA-256',encoder.encode(r.text)));
    const parts=[];
    for(let j=0;j<chunks.length;j++){
      const path=id+'/part-'+String(j+1).padStart(5,'0')+'.json';
      const payload={schema:'okto-chunk-v1',document_id:id,source_name:r.name,part_index:j+1,part_count:chunks.length,text:chunks[j]};
      const content=JSON.stringify(payload);if(bytes(content)>limit)throw Error('Часть превысила лимит: '+path);
      files.push({name:path,content});parts.push({path,index:j+1,text_bytes:bytes(chunks[j])});
    }
    let sourcePath=null;
    if(r.extension==='pdf'&&$('keepSource').checked){
      sourcePath=id+'/source.pdf';
      files.push({name:sourcePath,content:new Uint8Array(await r.sourceFile.arrayBuffer())});
    }
    const pageMap=r.pages?r.pages.map(p=>({page:p.page,marker:'===== СТРАНИЦА '+p.page+' =====',text_bytes:bytes(p.text)})):null;
    documents.push({document_id:id,source_name:r.name,source_type:r.extension,extraction_warning:r.warning||null,source_bytes:r.size,source_path:sourcePath,page_map:pageMap,text_bytes:bytes(r.text),text_sha256:sha,part_count:parts.length,parts});
  }
  const manifest={schema:'okto-package-v1',created_at:new Date().toISOString(),encoding:'UTF-8',split_strategy:'line-preferred-exact-text',pdf_pages:'page markers in text and page_map metadata',documents};
  files.unshift({name:'manifest.json',content:JSON.stringify(manifest,null,2)});
  return {blob:zipStore(files),parts:files.length-1};
}
drop.onclick=()=>input.click();drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click()}};
input.onchange=()=>{add([...input.files]);input.value=''};
drop.ondragover=e=>{e.preventDefault();drop.classList.add('over')};drop.ondragleave=()=>drop.classList.remove('over');
drop.ondrop=e=>{e.preventDefault();drop.classList.remove('over');add([...e.dataTransfer.files])};
$('clear').onclick=()=>{records.length=0;preview.textContent='Выберите документ.';render()};
exportBtn.onclick=async()=>{exportBtn.disabled=true;notice.textContent='Формируем ZIP…';try{const {blob,parts}=await makePackage();download('okto-package-'+new Date().toISOString().slice(0,10)+'.zip',blob);notice.textContent='ZIP сформирован: '+parts+' частей. Импорт в бот пока не реализован.'}catch(e){notice.textContent='Ошибка экспорта: '+(e.message||String(e))}finally{exportBtn.disabled=false}};
render();