/* Lee valores de la primera hoja sin construir un string XML de más de 512 MB.
 * Los originales no se modifican. ZIP: fflate 0.8.2, licencia MIT. */
(function(global){
 const decode=s=>s.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi,(_,k)=>k[0]==='#'?String.fromCodePoint(k[1].toLowerCase()==='x'?parseInt(k.slice(2),16):+k.slice(1)):({amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"}[k]));
 function* elements(bytes,tag){
  const decoder=new TextDecoder();let pending='';const close='</'+tag+'>';
  for(let pos=0;pos<bytes.length;pos+=65536){
   pending+=decoder.decode(bytes.subarray(pos,pos+65536),{stream:pos+65536<bytes.length});
   let end;
   while((end=pending.indexOf(close))!==-1){const start=pending.search(new RegExp('<'+tag+'(?:\\s|>)'));if(start<0)throw Error('XML inválido: '+tag);yield pending.slice(start,end+close.length);pending=pending.slice(end+close.length);}
  }
 }
 function attrs(xml){const a={};for(const m of xml.matchAll(/([\w:]+)="([^"]*)"/g))a[m[1]]=decode(m[2]);return a;}
 function leer(buffer){
  const files=fflate.unzipSync(new Uint8Array(buffer),{filter:f=>/^(xl\/sharedStrings.xml|xl\/workbook.xml|xl\/_rels\/workbook.xml.rels|xl\/worksheets\/sheet\d+.xml)$/.test(f.name)});
  const dec=new TextDecoder(),book=dec.decode(files['xl/workbook.xml']),first=book.match(/<sheet\s[^>]+/);
  if(!first)throw Error('El Excel no contiene hojas');
  const sheet=attrs(first[0]),rels=dec.decode(files['xl/_rels/workbook.xml.rels']);let target;
  for(const m of rels.matchAll(/<Relationship\s[^>]+/g)){const a=attrs(m[0]);if(a.Id===sheet['r:id'])target=a.Target;}
  const path=target?.startsWith('/')?target.slice(1):'xl/'+target;
  const bytes=files[path];if(!bytes)throw Error('No se encontró la primera hoja del Excel');
  const strings=[];
  if(files['xl/sharedStrings.xml'])for(const si of elements(files['xl/sharedStrings.xml'],'si')) strings.push([...si.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(m=>decode(m[1])).join(''));
  function* rows(){for(const xml of elements(bytes,'row')){
   const row=[];
   for(const m of xml.matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
    const a=attrs(m[1]);if(!a.r)continue;let col=0;for(const ch of a.r.match(/^[A-Z]+/)[0])col=col*26+ch.charCodeAt(0)-64;
    const content=m[2]||'',v=content.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/)?.[1];
    row[col-1]=a.t==='s'?strings[+v]:a.t==='inlineStr'?[...content.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x=>decode(x[1])).join(''):v===undefined?null:a.t==='str'||a.t==='e'?decode(v):Number(v);
   }
   yield row;
  }}
  return {SheetNames:[sheet.name],filasGrandes:rows()};
 }
 global.ExcelGrande={leer};
})(typeof self!=='undefined'?self:globalThis);
