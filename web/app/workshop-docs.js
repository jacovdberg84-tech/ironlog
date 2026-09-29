// IRONLOG/web/app/workshop-docs.js — Workshop library documents and uploads, Borris OEM.
// Part of the main app; index.html loads these files in order and they share one global scope.

// Internal workshop reference library. Download through authenticated fetch.
async function loadWorkshopDocuments() {
  const out = qs('workshopDocumentList');
  if (!out) return;
  out.textContent = 'Loading documents...';
  try {
    const data = await fetchJson(API + '/api/workshop/documents?q=' + encodeURIComponent(qs('workshopSearch')?.value || ''));
    out.replaceChildren();
    if (!data.documents.length) out.textContent = 'No matching documents. Upload your first manual above.';
    for (const doc of data.documents) {
      const card = document.createElement('div'); card.className = 'card';
      const title = document.createElement('strong'); title.textContent = doc.title;
      const details = document.createElement('p');
      details.textContent = [doc.doc_type,doc.manufacturer,doc.model,doc.revision && 'Revision ' + doc.revision,doc.applicability,(doc.size_bytes / 1048576).toFixed(1) + ' MB'].filter(Boolean).join(' · ');
      const button = document.createElement('button'); button.type = 'button'; button.textContent = 'Download PDF';
      button.addEventListener('click', async () => {
        button.disabled = true;
        try {
          const response = await fetch(API + '/api/workshop/documents/' + encodeURIComponent(doc.id) + '/file', {headers: authHeaders()});
          if (!response.ok) throw new Error('Download failed. Check your login and retry.');
          const url = URL.createObjectURL(await response.blob());
          const link = document.createElement('a'); link.href=url; link.download=doc.filename; document.body.appendChild(link); link.click(); link.remove();
          setTimeout(()=>URL.revokeObjectURL(url),60000);
        } catch(err) { qs('workshopUploadStatus').textContent=err.message; }
        finally { button.disabled=false; }
      });
      const indexStatus=document.createElement('p');indexStatus.textContent='Borris index: '+doc.index.status+(doc.index.pages?' · '+doc.index.pages+' pages':'')+(doc.index.message?' — '+doc.index.message:'');
      const indexButton=document.createElement('button');indexButton.type='button';indexButton.textContent='Index for Borris';
      indexButton.disabled=['queued','indexing'].includes(doc.index.status);
      indexButton.addEventListener('click',async()=>{
        indexButton.disabled=true;
        try {await fetchJson(API+'/api/workshop/documents/'+encodeURIComponent(doc.id)+'/index',{method:'POST',headers:authHeaders({'Content-Type':'application/json'}),body:'{}'});await loadWorkshopDocuments();}
        catch(err){qs('workshopUploadStatus').textContent=err.message;indexButton.disabled=false;}
      });
      card.append(title,details,indexStatus,indexButton,button); out.appendChild(card);
    }
    if (data.documents.length === data.limit) {
      const note=document.createElement('p'); note.textContent='Showing the first ' + data.limit + ' documents. Narrow your search to find older records.'; out.appendChild(note);
    }
  } catch(err) { out.textContent='Could not load documents: ' + err.message; }
}
function initWorkshopUploads() {
  qs('workshopAskForm')?.addEventListener('submit',async event=>{
    event.preventDefault();const form=event.currentTarget,button=form.querySelector('button'),out=qs('workshopAnswer');
    button.disabled=true;out.textContent='Searching indexed manuals...';
    try {const result=await fetchJson(API+'/api/workshop/ask',{method:'POST',headers:authHeaders({'Content-Type':'application/json'}),body:JSON.stringify({question:new FormData(form).get('question')})});out.textContent=result.short_answer;}
    catch(err){out.textContent='Could not answer: '+err.message;}finally{button.disabled=false;}
  });
  qs('workshopSearchForm')?.addEventListener('submit', event => {event.preventDefault(); loadWorkshopDocuments();});
  qs('workshopUploadForm')?.addEventListener('submit', async event => {
    event.preventDefault(); const form=event.currentTarget, button=form.querySelector('button[type="submit"]'), status=qs('workshopUploadStatus');
    const body=new FormData(form); const file=body.get('file');
    if (!file || file.size > 100*1024*1024 || !file.name.toLowerCase().endsWith('.pdf')) {status.textContent='Choose a PDF up to 100 MB.';return;}
    button.disabled=true;status.textContent='Uploading document...';
    try {
      const response=await fetch(API + '/api/workshop/documents', {method:'POST',headers:authHeaders(),body});
      const result=await response.json();if (!response.ok) throw new Error(result.error || 'Upload failed');
      form.reset();status.textContent='Document stored. Select Index for Borris to make its text searchable.';await loadWorkshopDocuments();
    } catch(err) {status.textContent='Upload failed: ' + err.message;}
    finally {button.disabled=false;}
  });
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',initWorkshopUploads); else initWorkshopUploads();

function initBorrisOem() {
  qs('borrisOemForm')?.addEventListener('submit',async event=>{
    event.preventDefault();const form=event.currentTarget,button=form.querySelector('button'),out=qs('borrisOemResult');
    button.disabled=true;out.textContent='Retrieving the approved source and checking indexed manuals...';
    try {
      const data=await fetchJson(API+'/api/oem/check',{method:'POST',headers:authHeaders({'Content-Type':'application/json'}),body:JSON.stringify(Object.fromEntries(new FormData(form)))});
      out.replaceChildren();
      const source=document.createElement('a');source.textContent=data.source.manufacturer+' — '+data.source.title;source.href=data.source.url;source.target='_blank';source.rel='noopener noreferrer';out.appendChild(source);
      const meta=document.createElement('p');meta.textContent='Retrieved '+data.source.retrieved_at+' · Revision: '+data.source.revision+'. '+data.source.scope;out.appendChild(meta);
      const comparison=document.createElement('p');comparison.style.whiteSpace='pre-wrap';comparison.textContent=data.comparison;out.appendChild(comparison);
      for(const [label,rows] of [['External source',data.external],['Internal manuals',data.internal]]) {
        const heading=document.createElement('h4');heading.textContent=label;out.appendChild(heading);
        if(!rows.length){const empty=document.createElement('p');empty.textContent='No matching passages found. Agreement or conflict cannot be established.';out.appendChild(empty);}
        for(const row of rows){const p=document.createElement('p');p.style.whiteSpace='pre-wrap';p.textContent='['+row.citation+'] '+(row.title||data.source.title)+(row.page?' — PDF page '+row.page:' — web page')+(row.revision?' · revision '+row.revision:'')+'\n'+row.excerpt;out.appendChild(p);}
      }
      const note=document.createElement('p');note.textContent=data.notice;out.appendChild(note);
    } catch(err){out.textContent=err.message;}finally{button.disabled=false;}
  });
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initBorrisOem);else initBorrisOem();
