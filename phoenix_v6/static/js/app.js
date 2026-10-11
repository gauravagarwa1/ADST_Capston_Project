function toggleSidebar(){document.getElementById('sidebar')?.classList.toggle('open')}
window.addEventListener('resize',()=>{if(window.innerWidth>1100) document.getElementById('sidebar')?.classList.remove('open')})
async function refreshNotificationCount(){try{const r=await fetch('/api/notifications');if(!r.ok)return;const d=await r.json();const el=document.getElementById('notificationCount');if(el){el.textContent=d.unread||0;el.style.display=(d.unread||0)?'inline-flex':'none'}}catch(e){}}
if(!window.matchMedia('(prefers-reduced-motion: reduce)').matches){document.addEventListener('DOMContentLoaded',()=>{document.querySelectorAll('.panel,.project-card,.project-detail-card,.portfolio-card,.kpi-card,.report-item').forEach((el,i)=>{el.animate([{opacity:0,transform:'translateY(8px)'},{opacity:1,transform:'translateY(0)'}],{duration:360,delay:Math.min(i*35,250),easing:'cubic-bezier(.2,.8,.2,1)',fill:'both'})})})}
document.addEventListener('DOMContentLoaded',refreshNotificationCount)
