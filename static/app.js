const $=x=>document.querySelector(x), start=$('#start'),range=$('#range'),status=$('#status'),results=$('#results');
function fmt(d){return new Intl.DateTimeFormat('sk-SK').format(d)}
function upd(){let d=new Date(start.value+'T12:00:00'),e=new Date(d);e.setDate(e.getDate()+6);range.textContent=`Programový týždeň: ${fmt(d)} – ${fmt(e)} (štvrtok → streda)`}
start.onchange=upd;upd();
function form(){let f=$('#excel').files[0];if(!f)throw Error('Najprv nahraj Excel.');let d=new Date(start.value+'T12:00:00');if(d.getDay()!=4)throw Error('Začiatok programového týždňa musí byť štvrtok.');let fd=new FormData();fd.append('excel',f);fd.append('cinema',$('#cinema').value);fd.append('start',start.value);return fd}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}
$('#preview').onclick=async()=>{try{status.className='status';status.textContent='Čítam Excel…';results.innerHTML='';let r=await fetch('/api/preview',{method:'POST',body:form()}),j=await r.json();if(!j.ok)throw Error(j.error);status.className='status ok';status.innerHTML=`🟢 Z Excelu som našla <b>${j.count}</b> predstavení.`;results.innerHTML=`<div class="tablewrap"><table><thead><tr><th>Dátum</th><th>Čas</th><th>Film</th><th>Sála</th><th>Atribút</th><th>Verzia</th></tr></thead><tbody>${j.expected.map(x=>`<tr><td>${esc(x.date)}</td><td>${esc(x.time)}</td><td>${esc(x.film)}</td><td>${esc(x.hall)}</td><td>${esc(x.attribute)}</td><td>${esc(x.version)}</td></tr>`).join('')}</tbody></table></div>`}catch(e){status.className='status bad';status.textContent='⚠️ '+e.message}}
$('#webpreview').onclick=async()=>{try{
  results.innerHTML='';
  let base=form();
  let all=[], diagnostics=[], byDay={}, expectedByDay={};
  for(let i=0;i<7;i++){
    let d=new Date(start.value+'T12:00:00'); d.setDate(d.getDate()+i);
    let day=d.toISOString().slice(0,10);
    status.className='status'; status.textContent=`Načítavam ${i+1}/7 – ${day}…`;
    let fd=form(); fd.append('day_index',String(i));
    let controller=new AbortController();
    let timer=setTimeout(()=>controller.abort(),35000);
    try{
      let r=await fetch('/api/web-preview-day',{method:'POST',body:fd,signal:controller.signal});
      let j=await r.json();
      if(!j.ok) throw Error(j.error);
      all.push(...(j.web||[])); byDay[day]=j.count||0; expectedByDay[day]=j.expected_count||0;
      diagnostics.push(j.diagnostic||`${day}: ${j.count||0} predstavení`);
    }catch(err){
      byDay[day]=0;
      diagnostics.push(`${day}: ${err.name==='AbortError'?'TIMEOUT po 35 s':err.message}`);
    }finally{clearTimeout(timer)}
  }
  let days=Object.keys(expectedByDay).sort().map(d=>`${d}: web ${byDay[d]||0} / Excel ${expectedByDay[d]||0}`).join(' • ');
  status.className='status ok'; status.innerHTML=`🟢 Diagnostika dokončená. Z Cinema City som prečítala <b>${all.length}</b> predstavení.<div class="small">${esc(days)}</div>`;
  results.innerHTML=`<div class="tablewrap"><table><thead><tr><th>Dátum</th><th>Čas</th><th>Film</th><th>Atribút</th><th>Verzia</th></tr></thead><tbody>${all.map(x=>`<tr><td>${esc(x.date)}</td><td>${esc(x.time)}</td><td>${esc(x.film)}</td><td>${esc(x.attribute)}</td><td>${esc(x.version)}</td></tr>`).join('')}</tbody></table></div><div class="small" style="margin-top:12px">Diagnostika: ${esc(diagnostics.join(' | '))}</div>`;
}catch(e){status.className='status bad';status.textContent='⚠️ '+e.message}}
$('#go').onclick=async()=>{try{status.className='status';status.textContent='Načítavam Cinema City pre 7 dní a porovnávam… môže to chvíľu trvať.';results.innerHTML='';let r=await fetch('/api/check',{method:'POST',body:form()}),j=await r.json();if(!j.ok)throw Error(j.error+(j.missing_days?.length?` Chýbajú dni: ${j.missing_days.join(', ')}.`:''));status.className='status '+(j.errors.length?'bad':'ok');status.innerHTML=j.errors.length?`🔴 Nájdených <b>${j.errors.length}</b> rozdielov. Excel: ${j.expected}, web: ${j.web}.`:`🟢 Program je správne nahodený. Skontrolovaných ${j.expected} predstavení.`;results.innerHTML=j.errors.map(x=>{let e=x.expected,w=x.web,t=x.type==='missing'?'Chýba na webe':x.type==='extra'?'Na webe navyše':x.type==='attribute'?'Nesedí atribút':'Nesedí dabing/titulky';let a=e||w;return `<div class="err"><b>${t}</b> — ${esc(a.date)} • ${esc(a.time)} • ${esc(a.film)}<div class="small">${e?`Excel: sála ${esc(e.hall)} | ${esc(e.attribute)} | ${esc(e.version)}`:''}${w?`<br>Web: ${esc(w.attribute)} | ${esc(w.version)}`:''}</div></div>`}).join('')}catch(e){status.className='status bad';status.textContent='⚠️ Kontrola nebola vykonaná: '+e.message}}
