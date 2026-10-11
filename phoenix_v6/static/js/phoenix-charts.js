(function(){
  const C=['#4F46E5','#0F9D8A','#E0527A','#7C3AED','#F59E0B','#0891B2','#DB2777','#16A34A'];
  const esc=v=>String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
  const num=v=>Number(v||0);
  function host(id){const el=document.getElementById(id); if(!el)return null; el.classList.add('phoenix-chart'); return el;}
  function empty(el,msg='No data available'){if(el)el.innerHTML=`<div class="chart-empty">${esc(msg)}</div>`;}
  function bar(id,data,title,opts={}){
    const el=host(id); if(!el)return;
    data=(data||[]).filter(x=>x&&x.label!==undefined).map(x=>({label:String(x.label),value:num(x.value)}));
    if(!data.length){empty(el);return;}
    const max=Math.max(...data.map(x=>x.value),1);
    const horizontal=data.length>=5 || opts.horizontal!==false;
    let html='<div class="chart-bars">';
    data.forEach((x,i)=>{
      const pct=Math.max(0,Math.min(100,(x.value/max)*100));
      html+=`<div class="chart-bar-row" title="${esc(x.label)}: ${x.value}"><div class="chart-bar-label">${esc(x.label)}</div><div class="chart-bar-track"><span style="--bar-width:${pct}%;width:${pct}%;background:${C[i%C.length]}"></span></div><div class="chart-bar-value">${Number.isInteger(x.value)?x.value:x.value.toFixed(1)}</div></div>`;
    });
    html+='</div>';
    el.innerHTML=html;
  }
  function donut(id,data){
    const el=host(id); if(!el)return;
    data=(data||[]).filter(x=>x&&x.label!==undefined).map(x=>({label:String(x.label),value:num(x.value)}));
    const total=data.reduce((a,b)=>a+b.value,0);
    if(!data.length||!total){empty(el);return;}
    let start=0, parts=[];
    data.forEach((x,i)=>{const p=x.value/total*100; parts.push(`${C[i%C.length]} ${start}% ${start+p}%`);start+=p;});
    el.innerHTML=`<div class="donut-wrap"><div class="donut-ring" style="background:conic-gradient(${parts.join(',')})"><div class="donut-hole"><b>${Math.round(total).toLocaleString()}</b><small>Total</small></div></div><div class="donut-legend">${data.map((x,i)=>`<div><i style="background:${C[i%C.length]}"></i><span>${esc(x.label)}</span><b>${x.value}</b></div>`).join('')}</div></div>`;
  }
  function scatter(id,data){
    const el=host(id); if(!el)return;
    data=(data||[]).filter(x=>Number.isFinite(num(x.x))&&Number.isFinite(num(x.y)));
    if(!data.length){empty(el);return;}
    const W=900,H=340,L=55,R=20,T=20,B=45;
    const sx=x=>L+(num(x)/100)*(W-L-R), sy=y=>H-B-(num(y)/100)*(H-T-B);
    let svg=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Accuracy versus completeness"><rect x="0" y="0" width="${W}" height="${H}" fill="transparent"/>`;
    [0,25,50,75,100].forEach(v=>{svg+=`<line x1="${sx(v)}" y1="${T}" x2="${sx(v)}" y2="${H-B}" stroke="#E2E8F0" stroke-width="1"/><line x1="${L}" y1="${sy(v)}" x2="${W-R}" y2="${sy(v)}" stroke="#E2E8F0" stroke-width="1"/><text x="${sx(v)}" y="${H-18}" fill="#94A3B8" font-size="12" text-anchor="middle">${v}</text><text x="${L-10}" y="${sy(v)+4}" fill="#94A3B8" font-size="12" text-anchor="end">${v}</text>`});
    data.forEach((p,i)=>{const r=Math.max(4,Math.min(10,Math.sqrt(num(p.size||1))));svg+=`<circle cx="${sx(p.x)}" cy="${sy(p.y)}" r="${r}" fill="#4F46E5" fill-opacity=".65"><title>Completeness: ${num(p.x)} · Accuracy: ${num(p.y)} · Score: ${num(p.size)}</title></circle>`});
    svg+=`<text x="${(L+W-R)/2}" y="${H-3}" fill="#64748B" font-size="12" text-anchor="middle">Completeness score</text><text x="14" y="${(T+H-B)/2}" fill="#64748B" font-size="12" text-anchor="middle" transform="rotate(-90 14 ${(T+H-B)/2})">Accuracy score</text></svg>`;
    el.innerHTML=svg;
  }
  window.PhoenixCharts={bar,donut,scatter,empty};
})();
