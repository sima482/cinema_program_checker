from flask import Flask, render_template, request, jsonify
from openpyxl import load_workbook
from datetime import datetime, timedelta
from io import BytesIO
import re, unicodedata
from playwright.sync_api import sync_playwright

app=Flask(__name__)
CINEMAS={'Eurovea':('eurovea','1012'),'Aupark':('aupark','1009'),'Polus':('polus','1011')}
DAY={'Mo':0,'Tu':1,'We':2,'Th':3,'Fr':4,'Sa':5,'Su':6}

def norm(s):
    s=unicodedata.normalize('NFKD',str(s or '')).encode('ascii','ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+',' ',s).strip()

def time_entries(v):
    if not v:return []
    txt=str(v).replace('\r','\n')
    times=re.findall(r'\b([0-2]?\d:[0-5]\d)\b',txt)
    only=re.search(r'only\s+([A-Za-z ]+)',txt,re.I)
    wo=re.search(r'w/o\s+([A-Za-z ]+)',txt,re.I)
    allowed=None
    if only: allowed={DAY[x] for x in only.group(1).split() if x in DAY}
    excluded={DAY[x] for x in wo.group(1).split() if x in DAY} if wo else set()
    return [(t,allowed,excluded) for t in times]

def expected_from_excel(data, cinema, start):
    wb=load_workbook(BytesIO(data),data_only=True)
    if cinema not in wb.sheetnames: raise ValueError(f'Hárok {cinema} v Exceli chýba.')
    ws=wb[cinema]; out=[]; hall=''
    for row in ws.iter_rows(min_row=2,values_only=True):
        if row[0]: hall=str(row[0]).strip()
        film=str(row[1] or '').strip()
        if not film: continue
        attr=str(row[4] or '').strip(); subdub=str(row[6] or '').strip()
        for cell in row[11:]:
            for tm,allowed,excluded in time_entries(cell):
                for i in range(7):
                    d=start+timedelta(days=i)
                    if allowed is not None and d.weekday() not in allowed: continue
                    if d.weekday() in excluded: continue
                    out.append({'date':d.isoformat(),'time':tm.zfill(5),'film':film,'hall':hall,'attribute':attr,'version':subdub})
    return out

def cinema_url(cinema,date):
    slug,cid=CINEMAS[cinema]
    return f'https://www.cinemacity.sk/cinemas/{slug}/{cid}#/buy-tickets-by-cinema?in-cinema={cid}&at={date}&view-mode=list'

def scrape_day(cinema,date):
    url=cinema_url(cinema,date)
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True,args=['--no-sandbox'])
        page=browser.new_page(locale='en-GB', user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36')
        page.goto(url,wait_until='domcontentloaded',timeout=60000)
        page.wait_for_timeout(5000)
        text=page.locator('body').inner_text()
        browser.close()
    # Parse the visible list. Each movie heading is followed by format/language/time lines.
    lines=[x.strip() for x in text.splitlines() if x.strip()]
    known_noise={'choose a date','all films','choose a movie','choose a screening type'}
    results=[]; current=None; attrs=[]; version=''
    attr_words=['2D','3D','VIP','4DX','DOLBY ATMOS','ATMOS','LASER BY BARCO','SUPER SCREEN','COMFORT','HFR','INFINITY VISION']
    for ln in lines:
        up=ln.upper()
        if re.fullmatch(r'[0-2]?\d:[0-5]\d',ln):
            if current: results.append({'date':date,'time':ln.zfill(5),'film':current,'attribute':' '.join(dict.fromkeys(attrs)),'version':version,'hall':''})
            continue
        if any(k in up for k in ['SUBTITLES:', 'DUBBED']) or re.search(r'\b(SLOVAK|CZECH)\b',up) and ('SUB' in up or 'DUB' in up):
            version=ln; continue
        matched=[a for a in attr_words if a in up]
        if matched:
            attrs.extend(matched); continue
        low=ln.lower()
        if low in known_noise or len(ln)>90 or ln.startswith(('Genre','Running time','Release date','Content category')): continue
        # movie titles are safest when followed later by a showtime; reset candidate on title-like lines
        if 1 < len(ln) < 70 and not re.search(r'\d+\s*mins?',ln,re.I) and not ln.startswith(('Image','×')):
            if not any(w in low for w in ['screening','cinema city','more information','suitable','violence','fear','profanity','action','drama','comedy','thriller','adventure','animation']):
                current=ln; attrs=[]; version=''
    return results,url

def compatible(e,w):
    return e['date']==w['date'] and e['time']==w['time'] and norm(e['film'])==norm(w['film'])

@app.get('/')
def home(): return render_template('index.html')
@app.get('/health')
def health(): return {'ok':True}
@app.post('/api/check')
def check():
    try:
        f=request.files.get('excel'); cinema=request.form.get('cinema','Eurovea'); start=datetime.strptime(request.form['start'],'%Y-%m-%d').date()
        exp=expected_from_excel(f.read(),cinema,start)
        web=[]; urls=[]
        for i in range(7):
            d=(start+timedelta(days=i)).isoformat(); r,u=scrape_day(cinema,d); web+=r; urls.append(u)
        if not web: return jsonify({'ok':False,'error':'Cinema City sa načítalo, ale nenašli sa žiadne predstavenia. Kontrola bola zastavená.'}),502
        errors=[]; used=set()
        for e in exp:
            idx=next((i for i,w in enumerate(web) if i not in used and compatible(e,w)),None)
            if idx is None: errors.append({'type':'missing','expected':e}); continue
            used.add(idx); w=web[idx]
            # Attribute/version comparison is deliberately explicit; hall often isn't exposed publicly.
            if e['attribute'] and norm(e['attribute']) not in norm(w['attribute']): errors.append({'type':'attribute','expected':e,'web':w})
            if e['version'] and norm(e['version']).replace('sub','subtitles').replace('dub','dubbed') not in norm(w['version']): errors.append({'type':'version','expected':e,'web':w})
        for i,w in enumerate(web):
            if i not in used: errors.append({'type':'extra','web':w})
        return jsonify({'ok':True,'expected':len(exp),'web':len(web),'errors':errors,'urls':urls})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500

if __name__=='__main__': app.run(host='0.0.0.0',port=5000,debug=True)
