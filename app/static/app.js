function showSection(id){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active-section'));
  const target=document.getElementById(id); if(target) target.classList.add('active-section');
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.target===id));
  window.scrollTo({top:0,behavior:'smooth'});
}
document.querySelectorAll('.nav-item').forEach(b=>b.addEventListener('click',()=>showSection(b.dataset.target)));

function brDate(iso){
  if(!iso)return '—';
  const d=new Date(iso); if(Number.isNaN(d.getTime())) return iso;
  return d.toLocaleString('pt-BR',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
const capture=document.getElementById('captureTime'); if(capture) capture.textContent=brDate(capture.dataset.iso);
document.querySelectorAll('.format-date').forEach(el=>el.textContent=brDate(el.dataset.iso));

const fileInput=document.getElementById('fileInput');
if(fileInput) fileInput.addEventListener('change',()=>{document.getElementById('fileName').textContent=fileInput.files[0]?.name || 'Nenhum arquivo selecionado'});

function filterTable(){
  const q=(document.getElementById('baseSearch')?.value||'').toLowerCase().trim();
  const st=document.getElementById('statusFilter')?.value||'';
  const imp=document.getElementById('implantFilter')?.value||'';
  let count=0;
  document.querySelectorAll('#basesTable tbody tr').forEach(tr=>{
    const show=(!q||tr.dataset.base.includes(q))&&(!st||tr.dataset.status===st)&&(!imp||tr.dataset.implant===imp);
    tr.style.display=show?'':'none'; if(show) count++;
  });
  const el=document.getElementById('visibleCount');if(el)el.textContent=count;
}

const params=new URLSearchParams(location.search); const toast=document.getElementById('toast');
if(toast&&params.has('import')){
  const ok=params.get('import')==='ok';
  toast.textContent=ok?'Planilha importada e novo snapshot registrado.':'Não foi possível importar a planilha. Confira o formato e a estrutura.';
  toast.classList.add('show',ok?'success':'error'); setTimeout(()=>toast.classList.remove('show'),4500);
  history.replaceState({},'',location.pathname);
}

if(window.DASH_DATA&&window.Chart){
  Chart.defaults.font.family='Inter, system-ui, sans-serif';
  Chart.defaults.color='#7b828c';
  const red='#d71938', green='#159b68', orange='#e79b25', gray='#cbd0d7', blue='#3478f6';
  const statusEntries=Object.entries(DASH_DATA.statuses);
  const statusEl=document.getElementById('statusChart');
  if(statusEl)new Chart(statusEl,{type:'doughnut',data:{labels:statusEntries.map(x=>x[0]),datasets:[{data:statusEntries.map(x=>x[1]),backgroundColor:statusEntries.map(([k])=>k==='EM FUNCIONAMENTO'?green:k.includes('PARCIAL')?orange:k.includes('DESATIV')?gray:red),borderWidth:0,hoverOffset:4}]},options:{maintainAspectRatio:false,cutout:'70%',plugins:{legend:{position:'right',labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:13,font:{size:9}}}}}});
  const d=Object.entries(DASH_DATA.day_bands); const daysEl=document.getElementById('daysChart');
  if(daysEl)new Chart(daysEl,{type:'bar',data:{labels:d.map(x=>x[0]),datasets:[{data:d.map(x=>x[1]),backgroundColor:d.map(([k])=>k==='Em implantação'?orange:k==='Sem dado'?gray:red),borderRadius:5,borderSkipped:false,maxBarThickness:36}]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false},ticks:{font:{size:9}}},y:{beginAtZero:true,grid:{color:'#f0f1f3'},ticks:{precision:0,font:{size:9}}}}}});
  const h=DASH_DATA.history||[]; const histEl=document.getElementById('historyChart');
  if(histEl)new Chart(histEl,{type:'line',data:{labels:h.map(x=>brDate(x.captured_at)),datasets:[{label:'% em funcionamento',data:h.map(x=>x.operational_rate),borderColor:red,backgroundColor:'rgba(215,25,56,.08)',fill:true,tension:.35,pointRadius:3,pointHoverRadius:5}]},options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false},ticks:{font:{size:8},maxRotation:0,autoSkip:true,maxTicksLimit:7}},y:{beginAtZero:true,max:100,grid:{color:'#f0f1f3'},ticks:{callback:v=>v+'%',font:{size:9}}}}}});
}
