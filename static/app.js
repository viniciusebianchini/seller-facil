const $ = id => document.getElementById(id);
let files = [], result = null, page = 0, busy = false;
function renderFiles(){
  $('file-list').replaceChildren();
  files.forEach((file,index)=>{
    const row=document.createElement('li'),badge=document.createElement('span'),name=document.createElement('span'),remove=document.createElement('button');
    badge.className='badge';badge.textContent=file.name.split('.').pop().toUpperCase();name.className='filename';name.textContent=file.name;
    remove.textContent='×';remove.setAttribute('aria-label','Remover '+file.name);remove.disabled=busy;remove.onclick=()=>{files.splice(index,1);renderFiles()};
    row.append(badge,name,remove);$('file-list').append(row);
  });
  $('process').disabled=busy||!files.some(f=>/\.xlsx\s*$/i.test(f.name))||!files.some(f=>/\.zip\s*$/i.test(f.name));
}
function addFiles(items){if(busy)return;if(result){files=[];result=null;$('results').hidden=true;$('empty').hidden=false;}for(const file of items){if(!/\.(xlsx|zip)\s*$/i.test(file.name)){$('error').textContent='Selecione apenas XLSX e ZIP originais da Shopee.';$('error').hidden=false;continue;}const index=files.findIndex(f=>f.name===file.name);if(index>=0)files[index]=file;else files.push(file)}renderFiles()}
$('files').onchange=e=>{addFiles(e.target.files);e.target.value=''};
for(const name of ['dragover','dragleave','drop'])$('drop').addEventListener(name,e=>{e.preventDefault();$('drop').classList.toggle('drag',name==='dragover');if(name==='drop')addFiles(e.dataTransfer.files)});
function showPage(){if(!result)return;$('label').src=result.previews[page];$('label').alt=`Etiqueta ${page+1} de ${result.etiquetas}`;$('page').textContent=`${page+1} de ${result.etiquetas} etiquetas`;$('previous').disabled=page===0;$('next').disabled=page===result.previews.length-1}
$('previous').onclick=()=>{page--;showPage()};$('next').onclick=()=>{page++;showPage()};
$('process').onclick=async()=>{
  busy=true;result=null;renderFiles();$('files').disabled=true;$('error').hidden=true;$('results').hidden=true;$('empty').hidden=false;$('status').textContent='Conferindo os pedidos e preparando os arquivos…';$('process').textContent='Preparando…';
  try{
    if(files.reduce((sum,f)=>sum+f.size,0)>80*1024*1024)throw Error('O limite total é 80 MB por envio.');
    const session=await fetch('/api/session').then(r=>r.json()),body=new FormData();files.forEach(f=>body.append('files',f));
    const response=await fetch('/api/process',{method:'POST',headers:{'X-App-Token':session.token},body});const data=await response.json();if(!response.ok)throw Error(data.error||'Não foi possível processar o lote.');
    result=data;page=0;$('empty').hidden=true;$('results').hidden=false;$('stats').replaceChildren();
    for(const [number,label] of [[data.pedidos,'pedidos'],[data.unidades,'unidades'],[data.skus,'SKUs']]){const cell=document.createElement('div'),value=document.createElement('strong');cell.className='stat';value.textContent=number;cell.append(value,document.createTextNode(label));$('stats').append(cell)}
    $('downloads').replaceChildren();for(const [key,label] of [['pdf','Baixar PDF'],['word','Lista em Word'],['arquivo','Baixar ZPL'],['zip','ZPL compactado'],['todos','Baixar todos os resultados']]){const a=document.createElement('a');a.href=data.downloads[key];a.textContent=label;a.download='';$('downloads').append(a)}
    showPage();$('status').textContent=`Lote pronto. ${data.etiquetas} etiquetas, incluindo ${data.alertas} de alerta de quantidade e ${data.separacao} de separação.`;
  }catch(error){$('error').textContent=error.message;$('error').hidden=false;$('status').textContent='Lote não gerado. Corrija os arquivos e tente novamente.'}
  finally{busy=false;$('files').disabled=false;$('process').textContent='Preparar etiquetas';renderFiles()}
};
if(document.modelContext?.registerTool){
  const lifecycle=new AbortController();
  window.addEventListener('pagehide',()=>lifecycle.abort(),{once:true});
  Promise.resolve(document.modelContext.registerTool({
    name:'consultar_lote',title:'Consultar lote preparado',description:'Retorna o resumo e os links do lote atualmente exibido, sem processar novos arquivos.',
    inputSchema:{type:'object',properties:{},additionalProperties:false},annotations:{readOnlyHint:true,untrustedContentHint:true},
    execute(input){if(!input||Object.keys(input).length)throw Error('Nenhum parâmetro é aceito.');return result?{pedidos:result.pedidos,etiquetas:result.etiquetas,unidades:result.unidades,downloads:result.downloads}:{status:busy?'processando':'nenhum lote preparado'}}
  },{signal:lifecycle.signal})).catch(()=>{});
}
