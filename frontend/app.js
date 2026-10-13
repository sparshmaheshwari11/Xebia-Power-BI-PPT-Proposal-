const form=document.getElementById("proposal-form");
const statusBox=document.getElementById("status");
const button=document.getElementById("submit");
form.addEventListener("submit",async(e)=>{
 e.preventDefault();
 const plain=(document.getElementById("requirements_text").value||"").trim();
 const upload=(document.getElementById("file").files||[]).length>0;
 if((plain.length>0 && upload) || (plain.length===0 && !upload)){statusBox.className="status show error";statusBox.textContent=plain.length>0&&upload?"Choose either plain text or a BRD upload, not both.":"Enter requirements in plain text or upload a BRD. Other fields are optional.";return;}
 statusBox.className="status show"; statusBox.textContent="Building the proposal: interpreting requirements, designing the solution narrative, and composing editable slides…"; button.disabled=true;
 try{
  const fd=new FormData(form);
  const res=await fetch("/api/generate",{method:"POST",body:fd});
  const payload=await res.json();
  if(!res.ok) throw new Error(payload.detail||"Proposal generation failed.");
  statusBox.className="status show success";
  statusBox.innerHTML=`<strong>Proposal created: ${payload.slide_count} slides</strong><br>${escapeHtml(payload.title)}<br><a href="${payload.download_url}">Download editable PowerPoint (.pptx)</a>`;
 }catch(err){statusBox.className="status show error";statusBox.textContent=err.message;}
 finally{button.disabled=false;}
});
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
