function prettyVersion(v){
  const raw=String(v||'').trim(); if(!raw) return '';
  const t=raw.toUpperCase().split(/\s+/);
  const mode=t.includes('SUB')?'SUB':t.includes('DUB')?'DUB':'';
  let lang='';
  if(t.some(x=>['SK','SVK'].includes(x))) lang='Slovak';
  else if(t.some(x=>['CZ','CS','CZE','CES'].includes(x))) lang='Czech';
  else if(t.some(x=>['EN','ENG'].includes(x))) lang='English';
  else if(t.some(x=>['HR','HRV'].includes(x))) lang='Croatian';
  else if(t.some(x=>['HU','HUN'].includes(x))) lang='Hungarian';
  else if(t.some(x=>['DE','GER','DEU'].includes(x))) lang='German';
  return [lang,mode].filter(Boolean).join(' ') || raw;
}
function prettyLanguage(v){
  const raw=String(v||'').trim(); if(!raw) return '';
  const x=raw.toUpperCase();
  return ({SK:'Slovak',SVK:'Slovak',CZ:'Czech',CS:'Czech',CZE:'Czech',EN:'English',ENG:'English',HR:'Croatian',HRV:'Croatian',HU:'Hungarian',HUN:'Hungarian',DE:'German',DEU:'German',FR:'French',IT:'Italian',ES:'Spanish',PL:'Polish'})[x] || raw;
}
const $=x=>document.querySelector(x), start=$('#start'),range=$('#range'),status=$('#status'),results=$('#results');
function fmt(d){return new Intl.DateTimeFormat('sk-SK').format(d)}
function upd(){let d=new Date(start.value+'T12:00:00'),e=new Date(d);e.setDate(e.getDate()+6);range.textContent=`Programový týždeň: ${fmt(d)} – ${fmt(e)} (štvrtok → streda)`}
start.onchange=upd;upd();
function form(){let f=$('#excel').files[0];if(!f)throw Error('Najprv nahraj Excel.');let d=new Date(start.value+'T12:00:00');if(d.getDay()!=4)throw Error('Začiatok programového týždňa musí byť štvrtok.');let fd=new FormData();fd.append('excel',f);fd.append('cinema',$('#cinema').value);fd.append('start',start.value);return fd}
function esc(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]))}
const names=['štvrtok','piatok','sobota','nedeľa','pondelok','utorok','streda'];
function selectedDayIndices(){return [...document.querySelectorAll('.day-check:checked')].map(x=>Number(x.value)).sort((a,b)=>a-b)}
function installDaySelector(){
  if(document.querySelector('#day-selector')) return;
  const box=document.createElement('div'); box.id='day-selector';
  box.style.cssText='margin-top:14px;padding:14px 16px;border:1px solid #285269;border-radius:12px;background:#0e2d3c';
  box.innerHTML=`<div style="font-weight:800;margin-bottom:9px">4. Dni na kontrolu</div><div style="display:flex;gap:14px;flex-wrap:wrap">${names.map((n,i)=>`<label style="display:flex;align-items:center;gap:6px;cursor:pointer"><input class="day-check" type="checkbox" value="${i}" checked> ${n}</label>`).join('')}</div><div class="small" style="margin-top:7px;opacity:.8">Označ iba dni, ktoré chceš kontrolovať.</div>`;
  range.insertAdjacentElement('afterend',box);
  const go=$('#go'); if(go) go.textContent='Skontrolovať vybrané dni';
}
installDaySelector();
function chosenDays(){const x=selectedDayIndices();if(!x.length)throw Error('Vyber aspoň jeden deň na kontrolu.');return x}
function dayISO(i){let d=new Date(start.value+'T12:00:00');d.setDate(d.getDate()+i);return d.toISOString().slice(0,10)}
function showProgress(prefix,i,pos,total){status.className='status';status.innerHTML=`⏳ ${esc(prefix)} <b>${names[i]} ${fmt(new Date(dayISO(i)+'T12:00:00'))}</b>… <span class="small">(${pos}/${total})</span>`}
function table(rows, hall=true){return `<div class="tablewrap"><table><thead><tr><th>Dátum</th><th>Čas</th><th>Film</th>${hall?'<th>Sála</th>':''}<th>Atribút</th><th>Pôvodný jazyk</th><th>Verzia</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.date)}</td><td>${esc(x.time)}</td><td>${esc(x.film)}</td>${hall?`<td>${esc(x.hall)}</td>`:''}<td>${esc(x.attribute)}</td><td>${esc(prettyLanguage(x.original_language))}</td><td>${esc(x.version)}</td></tr>`).join('')}</tbody></table></div>`}
function errorHtml(x){let e=x.expected,w=x.web,t=x.type==='missing'?'Chýba na webe':x.type==='extra'?'Na webe navyše':x.type==='time'?'Nesedí čas':x.type==='hall'?'Nesedí sála':x.type==='attribute'?'Nesedí atribút':x.type==='original_language'?'Nesedí pôvodný jazyk':'Nesedí dabing/titulky';let a=e||w;if(x.type==='time')return `<div class="err"><b>${t}</b> — ${esc(a.date)} • ${esc(a.film)}<div class="small">Excel: <b>${esc(e.time)}</b> • sála ${esc(e.hall)}<br>Cinema City: <b>${esc(w.time)}</b> • sála ${esc(w.hall)}</div></div>`;if(x.type==='original_language')return `<div class="err"><b>${t}</b> — ${esc(a.date)} • ${esc(a.time)} • ${esc(a.film)}<div class="small">Excel: <b>${esc(prettyLanguage(e?.original_language))}</b><br>Web: <b>${esc(prettyLanguage(w?.original_language))}</b></div></div>`;return `<div class="err"><b>${t}</b> — ${esc(a.date)} • ${esc(a.time)} • ${esc(a.film)}<div class="small">${e?`Excel: sála ${esc(e.hall)} | ${esc(e.attribute)} | ${esc(e.version)}`:''}${w?`<br>Web: sála ${esc(w.hall)} | ${esc(w.attribute)} | ${esc(prettyVersion(w.version))}`:''}</div></div>`}
$('#preview').onclick=async()=>{try{results.innerHTML='';status.className='status';status.textContent='Čítam Excel…';let r=await fetch('/api/preview',{method:'POST',body:form()}),j=await r.json();if(!j.ok)throw Error(j.error);let rows=[],days=chosenDays();for(let p=0;p<days.length;p++){let i=days[p];showProgress('Spracúvam Excel:',i,p+1,days.length);rows.push(...j.expected.filter(x=>x.date===dayISO(i)));await new Promise(res=>setTimeout(res,90));}status.className='status ok';status.innerHTML=`🟢 Z Excelu som pre vybrané dni našla <b>${rows.length}</b> predstavení.`;results.innerHTML=table(rows,true)}catch(e){status.className='status bad';status.textContent='⚠️ '+e.message}}
$('#webpreview').onclick=async()=>{try{results.innerHTML='';let all=[],byDay={},expectedByDay={},days=chosenDays();for(let p=0;p<days.length;p++){let i=days[p];showProgress('Načítavam Cinema City:',i,p+1,days.length);let fd=form();fd.append('day_index',String(i));let r=await fetch('/api/web-preview-day',{method:'POST',body:fd});let j=await r.json();if(!j.ok)throw Error(j.error);all.push(...(j.web||[]));byDay[j.day]=j.count||0;expectedByDay[j.day]=j.expected_count||0;}let summary=Object.keys(expectedByDay).sort().map(d=>`${d}: web ${byDay[d]||0} / Excel ${expectedByDay[d]||0}`).join(' • ');status.className='status ok';status.innerHTML=`🟢 Diagnostika dokončená. Z Cinema City som pre vybrané dni prečítala <b>${all.length}</b> predstavení.<div class="small">${esc(summary)}</div>`;results.innerHTML=table(all,true)}catch(e){status.className='status bad';status.textContent='⚠️ '+e.message}}
$('#go').onclick=async()=>{try{results.innerHTML='';let errors=[],totalE=0,totalW=0,days=chosenDays();for(let p=0;p<days.length;p++){let i=days[p];showProgress('Kontrolujem:',i,p+1,days.length);let fd=form();fd.append('day_index',String(i));let r=await fetch('/api/check-day',{method:'POST',body:fd});let j=await r.json();if(!j.ok)throw Error(j.error);totalE+=j.expected||0;totalW+=j.web||0;errors.push(...(j.errors||[]));window.__build=j.build||window.__build||'';}status.className='status '+(errors.length?'bad':'ok');status.innerHTML=errors.length?`🔴 Nájdených <b>${errors.length}</b> rozdielov. Excel: ${totalE}, web: ${totalW}. <span class="small">${window.__build||''}</span>`:`🟢 Program je správne nahodený. Skontrolovaných ${totalE} predstavení vo vybraných dňoch. <span class="small">${window.__build||''}</span>`;results.innerHTML=errors.map(errorHtml).join('')}catch(e){status.className='status bad';status.textContent='⚠️ Kontrola nebola vykonaná: '+e.message}}
