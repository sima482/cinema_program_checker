from flask import Flask, render_template, request, jsonify
from openpyxl import load_workbook
from datetime import datetime, timedelta, date
from io import BytesIO
import os, re, unicodedata
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

app = Flask(__name__)

# Current Cinema City SK cinema identifiers.
CINEMAS = {
    'Eurovea': ('eurovea', '1012'),
    'Aupark': ('aupark', '1010'),
    'Polus': ('polus', '1011'),
}
DAY = {'Mo': 0, 'Tu': 1, 'We': 2, 'Th': 3, 'Fr': 4, 'Sa': 5, 'Su': 6}
DAY_ALIASES = {
    0: ('mo', 'mon', 'monday', 'po', 'pondelok'),
    1: ('tu', 'tue', 'tuesday', 'ut', 'utorok'),
    2: ('we', 'wed', 'wednesday', 'st', 'streda'),
    3: ('th', 'thu', 'thursday', 'stv', 'štv', 'stvrtok', 'štvrtok'),
    4: ('fr', 'fri', 'friday', 'pi', 'piatok'),
    5: ('sa', 'sat', 'saturday', 'so', 'sobota'),
    6: ('su', 'sun', 'sunday', 'ne', 'nedela', 'nedeľa'),
}

ATTR_PATTERNS = [
    ('4DX 3D', ('4dx 3d',)),
    ('Super Screen 3D', ('super screen 3d', 'superscreen 3d')),
    ('Comfort 3D', ('comfort 3d',)),
    ('Atmos 3D', ('atmos 3d', 'dolby atmos 3d')),
    ('VIP 3D', ('vip 3d',)),
    ('3D', ('3d',)),
    ('4DX', ('4dx',)),
    ('Atmos', ('dolby atmos', 'atmos')),
    ('Comfort', ('comfort',)),
    ('Filmania', ('filmania', 'filmánia')),
    ('HFR', ('hfr', 'hrf')),
    ('Infinity Vision', ('infinity vision',)),
    ('Ladies Night', ('ladies night',)),
    ('Laser Barco', ('laser by barco', 'laserová projekcia barco', 'laser barco')),
    ('Marathon', ('marathon', 'maratón')),
    ('Special Event', ('special event', 'specia event')),
    ('Super Screen', ('super screen', 'superscreen')),
    ('VIP', ('vip',)),
]


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]+', ' ', s).strip()


def parse_clock(value):
    """Return HH:MM strings from Excel values, including real Excel time objects."""
    if value is None:
        return []
    if hasattr(value, 'hour') and hasattr(value, 'minute'):
        return [f'{value.hour:02d}:{value.minute:02d}']
    return [f'{int(h):02d}:{m}' for h, m in re.findall(r'\b([0-2]?\d):([0-5]\d)\b', str(value))]


def time_entries(v):
    if v is None or v == '':
        return []
    txt = str(v).replace('\r', '\n')
    times = parse_clock(v)
    only = re.search(r'only\s+([A-Za-z ]+)', txt, re.I)
    wo = re.search(r'w/o\s+([A-Za-z ]+)', txt, re.I)
    allowed = None
    if only:
        allowed = {DAY[x] for x in only.group(1).split() if x in DAY}
    excluded = {DAY[x] for x in wo.group(1).split() if x in DAY} if wo else set()
    return [(t, allowed, excluded) for t in times]


def header_weekday(value):
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.weekday()
    n = norm(value)
    for wd, aliases in DAY_ALIASES.items():
        for alias in aliases:
            if re.search(rf'(^|\s){re.escape(norm(alias))}(\s|$)', n):
                return wd
    return None


def expected_from_excel(data, cinema, start):
    """Read the CC SK schedule.

    Columns L onward are hourly placement buckets (9:00, 10:00, ...), not weekdays.
    A showtime without a day restriction applies to every day Thu-Wed.
    `only Su Sa` and `w/o Su Sa` restrict those individual showtimes.
    """
    wb = load_workbook(BytesIO(data), data_only=True)
    sheet = next((s for s in wb.sheetnames if norm(s) == norm(cinema)), None)
    if not sheet:
        raise ValueError(f'Hárok {cinema} v Exceli chýba.')
    ws = wb[sheet]
    out, hall = [], ''

    for row_idx in range(2, ws.max_row + 1):
        row = [ws.cell(row_idx, c).value for c in range(1, ws.max_column + 1)]

        # The schedule is followed by a separate Legend / Features IN / Features OUT
        # table.  It also contains durations such as 01:53, so it must never be
        # interpreted as screening times.
        if norm(row[1]) == 'legend' or norm(row[7]) == 'features in':
            break

        if row[0] not in (None, ''):
            hall = str(row[0]).strip()
        film = str(row[1] or '').strip()
        if not film:
            continue
        original = str(row[2] or '').strip()
        attr = str(row[4] or '').strip()
        subdub = str(row[6] or '').strip()

        # L onward are time buckets. Each populated cell may contain an actual
        # showtime plus an optional day rule such as `only Su Sa` / `w/o Su Sa`.
        for value in row[11:]:
            for tm, allowed, excluded in time_entries(value):
                for i in range(7):
                    target = start + timedelta(days=i)
                    if allowed is not None and target.weekday() not in allowed:
                        continue
                    if target.weekday() in excluded:
                        continue
                    out.append({
                        'date': target.isoformat(), 'time': tm, 'film': film,
                        'original': original, 'hall': hall,
                        'attribute': attr, 'version': subdub,
                    })
    return out

def cinema_url(cinema, day):
    slug, cid = CINEMAS[cinema]
    return f'https://www.cinemacity.sk/cinemas/{slug}/{cid}#/buy-tickets-by-cinema?in-cinema={cid}&at={day}&view-mode=list'


def canon_attributes(text):
    n = norm(text)
    found = []
    for canonical, variants in ATTR_PATTERNS:
        if any(norm(v) in n for v in variants):
            # Do not add 3D separately when it is already represented by a combined format.
            if canonical == '3D' and any(x.endswith('3D') for x in found):
                continue
            if canonical not in found:
                found.append(canonical)
    return found


def canon_version(text):
    n = norm(text)
    if not n:
        return ''
    mode = ''
    if any(x in n for x in ('titulky', 'subtitle', 'subtitles', ' sub ')):
        mode = 'SUB'
    elif any(x in n for x in ('dabing', 'dubbed', ' dubbing', ' dub ')):
        mode = 'DUB'
    languages = []
    for canonical, variants in {
        'SK': ('slovencina', 'slovak', 'sk'),
        'CZ': ('cestina', 'czech', 'cz'),
        'EN': ('anglictina', 'english', 'en'),
        'HU': ('madarcina', 'hungarian', 'hu'),
        'DE': ('nemcina', 'german', 'de'),
    }.items():
        if any(re.search(rf'(^|\s){re.escape(norm(v))}(\s|$)', n) for v in variants):
            languages.append(canonical)
    # Cinema City strings include original language first, e.g. English (Subtitles: Slovak).
    # For checking we need the dubbed/subtitle language, which is normally the last language.
    lang = languages[-1] if languages else ''
    return ' '.join(x for x in (mode, lang) if x)


def excel_version(text):
    n = norm(text)
    mode = 'SUB' if ('sub' in n or 'tit' in n) else ('DUB' if ('dub' in n or 'dab' in n) else '')
    lang = ''
    lang_map = {
        'SK': ('sk', 'svk', 'slovak', 'slovencina'),
        'CZ': ('cz', 'cze', 'czech', 'cestina'),
        'EN': ('en', 'eng', 'english', 'anglictina'),
        'HU': ('hu', 'hun', 'hungarian', 'madarcina'),
        'DE': ('de', 'ger', 'german', 'nemcina'),
    }
    for canonical, variants in lang_map.items():
        if any(re.search(rf'(^|\s){re.escape(norm(v))}(\s|$)', n) for v in variants):
            lang = canonical
            break
    return ' '.join(x for x in (mode, lang) if x)


def known_titles_from_expected(expected):
    """Use the Excel titles as anchors when reading Cinema City.

    This is much safer than guessing movie headings from every visible label on the site.
    Both the local and original title are accepted.
    """
    titles = []
    for item in expected or []:
        for key in ('film', 'original'):
            title = re.sub(r'\s+', ' ', str(item.get(key) or '')).strip()
            if title and norm(title) not in {norm(x) for x in titles}:
                titles.append(title)
    return titles


def match_known_title(line, movie_titles):
    """Return the Excel title represented by a Cinema City heading, if any."""
    ln = norm(line)
    if not ln:
        return None
    # Prefer the longest title so one title cannot steal another title's prefix.
    for title in sorted(movie_titles, key=lambda x: len(norm(x)), reverse=True):
        tn = norm(title)
        if ln == tn:
            return title
        # Cinema City sometimes appends a short format marker to the heading
        # (for example IV = Infinity Vision). Accept only very small known suffixes.
        if ln.startswith(tn + ' '):
            suffix = ln[len(tn):].strip()
            if suffix in {'iv', 'vip', '3d', '4dx', '4dx 3d', 'hfr'}:
                return title
    return None


def parse_page_text(text, movie_titles, day):
    lines = [re.sub(r'\s+', ' ', x).strip() for x in text.splitlines() if x.strip()]
    results = []
    current = None
    attrs = []
    version_raw = ''
    seen_screening_data = False

    for ln in lines:
        low = norm(ln)
        if low.startswith(('o eurovea bratislava', 'o aupark bratislava', 'o polus bratislava',
                           'about eurovea bratislava', 'about aupark bratislava', 'about polus bratislava')):
            break

        matched_title = match_known_title(ln, movie_titles)
        if matched_title:
            current = matched_title
            attrs, version_raw, seen_screening_data = [], '', False
            # A short suffix on a heading can itself be a format marker.
            suffix = norm(ln)[len(norm(matched_title)):].strip()
            if suffix == 'iv':
                attrs = ['Infinity Vision']
                seen_screening_data = True
            elif suffix:
                for a in canon_attributes(suffix):
                    if a not in attrs:
                        attrs.append(a)
                seen_screening_data = bool(attrs)
            continue
        if not current:
            continue

        # Format and language can be on separate lines or concatenated, e.g.
        # "2DFRANCÚZŠTINA (DABING: SLOVENČINA)".
        line_attrs = canon_attributes(ln)
        if line_attrs:
            for a in line_attrs:
                if a not in attrs:
                    attrs.append(a)
            seen_screening_data = True

        cv = canon_version(ln)
        if cv:
            version_raw = ln
            seen_screening_data = True

        times = parse_clock(ln)
        if times and (seen_screening_data or current):
            for tm in times:
                results.append({
                    'date': day, 'time': tm, 'film': current,
                    'attribute': ' '.join(attrs),
                    'version': version_raw,
                    'hall': '',
                })
            # Do not leak one screening variant into the next time group.
            attrs, version_raw, seen_screening_data = [], '', False
    return results


def scrape_day_safe(cinema, day, movie_titles):
    """Load one day with strict time limits so a bad Cinema City page cannot hang forever."""
    url = cinema_url(cinema, day)
    results = []
    diagnostic = ''
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--no-sandbox', '--disable-dev-shm-usage'])
        try:
            context = browser.new_context(
                locale='sk-SK',
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
                viewport={'width': 1440, 'height': 1100},
            )
            page = context.new_page()
            page.set_default_timeout(7000)
            try:
                response = page.goto(url, wait_until='domcontentloaded', timeout=15000)
                status = response.status if response else None
                if status and status >= 400:
                    diagnostic = f'{day}: HTTP {status}'
                else:
                    # Never wait indefinitely for Cinema City's SPA. If showtimes do not
                    # appear quickly, still inspect whatever text is available.
                    try:
                        page.wait_for_function("() => /\b[0-2]?\d:[0-5]\d\b/.test(document.body.innerText)", timeout=7000)
                    except PlaywrightTimeoutError:
                        pass
                    page.wait_for_timeout(900)
                    try:
                        text = page.locator('body').inner_text(timeout=5000)
                    except PlaywrightTimeoutError:
                        text = ''
                    results = parse_page_text(text, movie_titles, day) if text else []
                    diagnostic = f'{day}: {len(results)} predstavení'
            except PlaywrightTimeoutError:
                diagnostic = f'{day}: TIMEOUT – Cinema City neodpovedalo včas'
            except Exception as exc:
                diagnostic = f'{day}: chyba načítania ({type(exc).__name__})'
            finally:
                try:
                    page.close()
                except Exception:
                    pass
                try:
                    context.close()
                except Exception:
                    pass
        finally:
            try:
                browser.close()
            except Exception:
                pass
    return results, url, diagnostic


def scrape_week(cinema, start, movie_titles):
    """Read seven days with a hard per-day timeout and continue after failures."""
    all_results, urls, diagnostics = [], [], []
    for i in range(7):
        day = (start + timedelta(days=i)).isoformat()
        parsed, url, diagnostic = scrape_day_safe(cinema, day, movie_titles)
        urls.append(url)
        diagnostics.append(diagnostic)
        all_results.extend(parsed)
    return all_results, urls, diagnostics

def same_title(expected, web):
    candidates = [expected.get('film', ''), expected.get('original', '')]
    wn = norm(web.get('film', ''))
    return any(norm(c) == wn for c in candidates if c)


def compatible(e, w):
    return e['date'] == w['date'] and e['time'] == w['time'] and same_title(e, w)


def attributes_match(expected, actual):
    e = set(canon_attributes(expected))
    w = set(canon_attributes(actual))
    return not e or e.issubset(w)


def version_match(expected, actual):
    e = excel_version(expected)
    w = canon_version(actual)
    return not e or e == w


@app.get('/')
def home():
    return render_template('index.html')


@app.get('/health')
def health():
    return {'ok': True}


@app.post('/api/preview')
def preview():
    try:
        f = request.files.get('excel')
        if not f:
            raise ValueError('Najprv nahraj Excel.')
        cinema = request.form.get('cinema', 'Eurovea')
        start = datetime.strptime(request.form['start'], '%Y-%m-%d').date()
        if start.weekday() != 3:
            raise ValueError('Začiatok programového týždňa musí byť štvrtok.')
        exp = expected_from_excel(f.read(), cinema, start)
        return jsonify({'ok': True, 'expected': exp, 'count': len(exp)})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 400


@app.post('/api/web-preview-day')
def web_preview_day():
    """V6 diagnostic endpoint: one HTTP request = one Cinema City day.

    This lets the browser show real progress and prevents one stuck day from
    hiding the other six days.
    """
    try:
        cinema = request.form.get('cinema', 'Eurovea')
        if cinema not in CINEMAS:
            raise ValueError('Neznáme kino.')
        start = datetime.strptime(request.form['start'], '%Y-%m-%d').date()
        if start.weekday() != 3:
            raise ValueError('Začiatok programového týždňa musí byť štvrtok.')
        day_index = int(request.form.get('day_index', '0'))
        if day_index < 0 or day_index > 6:
            raise ValueError('Neplatný deň diagnostiky.')
        f = request.files.get('excel')
        if not f:
            raise ValueError('Najprv nahraj Excel.')
        exp = expected_from_excel(f.read(), cinema, start)
        titles = known_titles_from_expected(exp)
        day = (start + timedelta(days=day_index)).isoformat()
        web, url, diagnostic = scrape_day_safe(cinema, day, titles)
        expected_count = sum(1 for x in exp if x['date'] == day)
        return jsonify({
            'ok': True, 'day': day, 'web': web, 'count': len(web),
            'expected_count': expected_count, 'diagnostic': diagnostic, 'url': url,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.post('/api/web-preview')
def web_preview():
    try:
        cinema = request.form.get('cinema', 'Eurovea')
        if cinema not in CINEMAS:
            raise ValueError('Neznáme kino.')
        start = datetime.strptime(request.form['start'], '%Y-%m-%d').date()
        if start.weekday() != 3:
            raise ValueError('Začiatok programového týždňa musí byť štvrtok.')
        f = request.files.get('excel')
        if not f:
            raise ValueError('Najprv nahraj Excel – vo V5 ho používam aj na bezpečné rozpoznanie názvov filmov na Cinema City.')
        exp = expected_from_excel(f.read(), cinema, start)
        titles = known_titles_from_expected(exp)
        web, urls, diagnostics = scrape_week(cinema, start, titles)
        by_day, expected_by_day = {}, {}
        for x in web:
            by_day[x['date']] = by_day.get(x['date'], 0) + 1
        for x in exp:
            expected_by_day[x['date']] = expected_by_day.get(x['date'], 0) + 1
        return jsonify({
            'ok': True, 'web': web, 'count': len(web), 'by_day': by_day,
            'expected_by_day': expected_by_day,
            'diagnostics': diagnostics, 'urls': urls,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.post('/api/check')
def check():
    try:
        f = request.files.get('excel')
        if not f:
            raise ValueError('Najprv nahraj Excel.')
        cinema = request.form.get('cinema', 'Eurovea')
        if cinema not in CINEMAS:
            raise ValueError('Neznáme kino.')
        start = datetime.strptime(request.form['start'], '%Y-%m-%d').date()
        if start.weekday() != 3:
            raise ValueError('Začiatok programového týždňa musí byť štvrtok.')

        exp = expected_from_excel(f.read(), cinema, start)
        titles = known_titles_from_expected(exp)
        web, urls, diagnostics = scrape_week(cinema, start, titles)
        days_with_data = {x['date'] for x in web}

        # Never produce a giant false-error report when Cinema City failed to load.
        if not web:
            return jsonify({
                'ok': False,
                'error': 'Cinema City sa nepodarilo načítať alebo stránka nevrátila žiadne predstavenia. Kontrola bola bezpečne zastavená – nič neoznačujem ako chýbajúce.',
                'diagnostics': diagnostics,
            }), 502

        expected_days = {x['date'] for x in exp}
        missing_loaded_days = sorted(d for d in expected_days if d not in days_with_data)
        if missing_loaded_days:
            return jsonify({
                'ok': False,
                'error': 'Cinema City neposkytlo dáta pre všetky dni, v ktorých Excel obsahuje predstavenia. Kontrola bola zastavená, aby nevznikli falošné chyby.',
                'missing_days': missing_loaded_days,
                'diagnostics': diagnostics,
            }), 502

        errors, used = [], set()
        for e in exp:
            idx = next((i for i, w in enumerate(web) if i not in used and compatible(e, w)), None)
            if idx is None:
                errors.append({'type': 'missing', 'expected': e})
                continue
            used.add(idx)
            w = web[idx]
            if not attributes_match(e['attribute'], w['attribute']):
                errors.append({'type': 'attribute', 'expected': e, 'web': w})
            if not version_match(e['version'], w['version']):
                errors.append({'type': 'version', 'expected': e, 'web': w})

        for i, w in enumerate(web):
            if i not in used:
                errors.append({'type': 'extra', 'web': w})

        return jsonify({
            'ok': True, 'expected': len(exp), 'web': len(web), 'errors': errors,
            'urls': urls, 'diagnostics': diagnostics,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '5000')), debug=True)
