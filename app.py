from flask import Flask, render_template, request, jsonify
from openpyxl import load_workbook
from datetime import datetime, timedelta, date
from io import BytesIO
import os, re, unicodedata, requests

app = Flask(__name__)
BUILD_VERSION = 'V35'

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
    ('Infinity Vision', ('infinity vision', 'infinity')),
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
        original_language = str(row[5] or '').strip()
        subdub = str(row[6] or '').strip()
        # Canonical structured values used by the comparison.
        excel_original_code = _language_code(original_language)
        excel_loc_mode, excel_loc_lang = _version_parts(subdub, excel=True)

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
                        'attribute': attr, 'original_language': original_language, 'version': subdub,
                        'original_language_code': excel_original_code,
                        'localization_mode': excel_loc_mode, 'localization_language': excel_loc_lang,
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
    # Match SUB/DUB also when it is the first token, e.g. Quickbook gives
    # `SUB SK` while Excel gives `Slovak SUB`. Both must canonicalize to SUB SK.
    tokens = set(n.split())
    if ('sub' in tokens or 'titulky' in n or 'subtitle' in n or 'subtitles' in n):
        mode = 'SUB'
    elif ('dub' in tokens or 'dabing' in n or 'dubbed' in n or 'dubbing' in n):
        mode = 'DUB'
    languages = []
    for canonical, variants in {
        'SK': ('slovencina', 'slovak', 'sk'),
        'CZ': ('cestina', 'czech', 'cz', 'cs'),
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
        'CZ': ('cz', 'cs', 'cze', 'ces', 'czech', 'cestina'),
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


def _flatten_values(obj):
    """Collect descriptive scalar values from a Quickbook film/event object."""
    out = []
    if isinstance(obj, dict):
        for value in obj.values():
            out.extend(_flatten_values(value))
    elif isinstance(obj, (list, tuple)):
        for value in obj:
            out.extend(_flatten_values(value))
    elif isinstance(obj, (str, int, float)) and not isinstance(obj, bool):
        out.append(str(obj))
    return out


def _clean_hall(event):
    """Cinema City Quickbook uses auditoriumTinyName for the public hall label."""
    value = (event.get('auditoriumTinyName') or event.get('auditoriumName') or
             event.get('auditorium') or event.get('screenName') or event.get('screen') or '')
    return re.sub(r'\s+', ' ', str(value)).strip()


def _clean_version_from_quickbook(event, film):
    """Return only SUB/DUB + target language; never expose the raw API blob."""
    # Prefer dedicated fields if Cinema City supplies them.
    parts = []
    keys = ('presentationMethodAndLanguage', 'languageVersion', 'eventLanguage',
            'filmLanguage', 'subtitleLanguage', 'subtitlesLanguage', 'dubLanguage',
            'dubbingLanguage', 'version', 'language')
    for obj in (event, film):
        for key in keys:
            value = obj.get(key)
            if value not in (None, '', [], {}):
                parts.extend(_flatten_values(value))
    dedicated = ' '.join(parts)
    cv = canon_version(dedicated)
    if cv:
        return cv

    # Some Quickbook installations encode the presentation in descriptive
    # event attributes instead of a dedicated language property.  Looking at
    # the event only avoids accidentally inheriting a movie-wide format.
    event_text = ' '.join(_flatten_values(event))
    return canon_version(event_text)


def _language_code(text):
    """Normalize a language label/code to SK/CZ/EN/HU/DE/HR when known."""
    n = norm(text)
    tokens = set(n.split())
    mapping = {
        'SK': ('sk','svk','slovak','slovencina'),
        'CZ': ('cz','cs','cze','ces','czech','cestina'),
        'EN': ('en','eng','english','anglictina'),
        'HU': ('hu','hun','hungarian','madarcina'),
        'DE': ('de','ger','deu','german','nemcina'),
        'HR': ('hr','hrv','croatian','chorvatcina'),
    }
    for code, variants in mapping.items():
        if any(norm(v) in tokens for v in variants):
            return code
    return ''


def _quickbook_language_data(event, film):
    """Return (spoken/original language, localisation mode, localisation language).

    Cinema City displays screenings as e.g. `ČEŠTINA` or
    `ANGLIČTINA (TITULKY: SLOVENČINA)`.  The language outside brackets is
    the screening/original language; the bracket supplies SUB/DUB target.
    Quickbook may expose that first language through original-lang-* OR its
    structured `languages`/language fields, so do not default a missing value
    to Slovak.
    """
    attr_ids = [norm(x).replace(' ', '-') for x in (event.get('attributeIds') or [])]
    film_attr_ids = [norm(x).replace(' ', '-') for x in (film.get('attributeIds') or [])]
    all_ids = attr_ids + film_attr_ids

    mode = 'SUB' if 'subbed' in attr_ids else ('DUB' if 'dubbed' in attr_ids else '')
    target = ''
    if mode:
        prefixes = ('first-subbed-lang-', 'subbed-lang-') if mode == 'SUB' else ('first-dubbed-lang-', 'dubbed-lang-')
        for aid in attr_ids:
            for prefix in prefixes:
                if aid.startswith(prefix):
                    raw = aid[len(prefix):]
                    target = _language_code(raw) or raw.upper()
                    break
            if target:
                break

    original = ''
    for aid in all_ids:
        for prefix in ('original-lang-', 'original-language-'):
            if aid.startswith(prefix):
                original = _language_code(aid[len(prefix):]) or aid[len(prefix):].upper()
                break
        if original:
            break

    # Some SK Quickbook events (notably screenings shown simply as `ČEŠTINA`)
    # omit original-lang-* but expose the spoken language in structured fields.
    # Collect codes in source order. For SUB/DUB, exclude the known target; the
    # remaining language is the language outside Cinema City's brackets.
    if not original:
        candidates = []
        for obj, keys in ((event, ('languages','eventLanguage','filmLanguage','language','presentationMethodAndLanguage')),
                          (film, ('languages','originalLanguage','filmLanguage','language'))):
            for key in keys:
                value = obj.get(key)
                if value in (None, '', [], {}):
                    continue
                for raw in _flatten_values(value):
                    code = _language_code(raw)
                    if code and code not in candidates:
                        candidates.append(code)
        if mode and target:
            non_target = [c for c in candidates if c != target]
            original = non_target[0] if non_target else ''
        elif candidates:
            original = candidates[0]

    return original, mode, target


def localization_key(text, excel=False):
    """Canonical G-column/localisation identity. Bare Excel SUB/DUB means Slovak."""
    mode, lang = _version_parts(text, excel=excel)
    return mode, lang


def original_language_match(expected, actual):
    e = _language_code(expected)
    w = _language_code(actual)
    if not e or not w:
        return True  # do not invent a mismatch when Quickbook lacks the datum
    return e == w

def _quickbook_event_to_web(event, film, requested_day):
    raw_dt = str(event.get('eventDateTime') or event.get('dateTime') or event.get('startTime') or '')
    day, tm = requested_day, ''
    if raw_dt:
        try:
            dt = datetime.fromisoformat(raw_dt.replace('Z', '+00:00'))
            day, tm = dt.date().isoformat(), dt.strftime('%H:%M')
        except ValueError:
            m = re.search(r'(20\d{2}-\d{2}-\d{2})[T ]([0-2]\d:[0-5]\d)', raw_dt)
            if m:
                day, tm = m.group(1), m.group(2)
    if not tm:
        raw_time = str(event.get('time') or event.get('start') or '')
        m = re.search(r'([0-2]?\d):([0-5]\d)', raw_time)
        if m:
            tm = f'{int(m.group(1)):02d}:{m.group(2)}'

    film_name = str(film.get('name') or film.get('title') or event.get('filmName') or event.get('name') or '').strip()
    hall = _clean_hall(event)

    # Screening attributes belong to the event.  Film metadata contains genre,
    # age rating and other words that must not become screening attributes.
    event_blob = ' '.join(_flatten_values(event))
    attrs = canon_attributes(event_blob)
    original_language, mode, target = _quickbook_language_data(event, film)
    # If there is no bracket/localisation, show the actual language Cinema City
    # gives for the screening (e.g. CZECH), never an invented Slovak default.
    version = ' '.join(x for x in (target, mode) if x) if mode else original_language

    return {
        'date': day, 'time': tm, 'film': film_name, 'attribute': ' '.join(attrs),
        'version': version, 'original_language': original_language, 'hall': hall,
        'original_language_code': _language_code(original_language) or original_language,
        'localization_mode': mode, 'localization_language': target,
        # Keep Cinema City's structured language/presentation metadata internally.
        # It is used below to learn the meaning of SK Quickbook presentation codes
        # from the whole week instead of guessing that the original film language
        # (for example EN) is the subtitle/dubbing language.
        '_presentation_code': str(event.get('presentationCode') or '').strip(),
        '_languages': event.get('languages'),
    }


QUICKBOOK_GROUP_CANDIDATES = [str(x) for x in range(10100, 10121)]


def _qb_get(base, path, timeout=12):
    url = f'{base}{path}'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131 Safari/537.36',
        'Accept': 'application/json, text/plain, */*',
        'Referer': 'https://www.cinemacity.sk/',
    }
    r = requests.get(url, headers=headers, timeout=timeout)
    if r.status_code != 200:
        return None, f'HTTP {r.status_code}', url
    try:
        payload = r.json()
    except ValueError:
        return None, 'odpoveď nie je JSON', url
    body = payload.get('body', payload) if isinstance(payload, dict) else payload
    return body, '', url


def discover_quickbook_group(cinema_id, horizon):
    """Find the SK Quickbook group instead of guessing a country group id."""
    diagnostics = []
    # Cinema City CZ publicly uses the same data-api-service shape with group 10101.
    # SK's group is discovered by asking which candidate recognizes the SK cinema id.
    for group in QUICKBOOK_GROUP_CANDIDATES:
        base = f'https://www.cinemacity.sk/sk/data-api-service/v1/quickbook/{group}'
        body, err, _ = _qb_get(base, f'/dates/in-cinema/{cinema_id}/until/{horizon}?attr=&lang=sk_SK', timeout=6)
        if err:
            continue
        dates = body.get('dates', []) if isinstance(body, dict) else []
        if dates:
            diagnostics.append(f'Quickbook SK group {group}: kino {cinema_id} rozpoznané, {len(dates)} dostupných dní')
            return group, base, diagnostics
        # A valid group may have no dates on a far horizon; verify through cinema list.
        body2, err2, _ = _qb_get(base, f'/cinemas/with-event/until/{horizon}?attr=&lang=sk_SK', timeout=6)
        if not err2 and isinstance(body2, dict):
            cinemas = body2.get('cinemas', [])
            if any(str(c.get('id')) == str(cinema_id) for c in cinemas):
                diagnostics.append(f'Quickbook SK group {group}: kino {cinema_id} nájdené v zozname kín')
                return group, base, diagnostics
    diagnostics.append('Quickbook SK: nepodarilo sa nájsť group ID, ktoré pozná vybrané kino')
    return None, None, diagnostics


def fetch_quickbook_week(cinema, start):
    """V8: fetch exact Thu-Wed schedule from Cinema City's public Quickbook JSON API."""
    _, cid = CINEMAS[cinema]
    end = start + timedelta(days=6)
    horizon = (end + timedelta(days=1)).isoformat()
    group, base, diagnostics = discover_quickbook_group(cid, horizon)
    if not base:
        return [], [], diagnostics

    web, urls = [], []
    for i in range(7):
        day = (start + timedelta(days=i)).isoformat()
        body, err, url = _qb_get(base, f'/film-events/in-cinema/{cid}/at-date/{day}?attr=&lang=sk_SK', timeout=20)
        urls.append(url)
        if err:
            diagnostics.append(f'{day}: {err}')
            continue
        if not isinstance(body, dict):
            diagnostics.append(f'{day}: neznámy formát odpovede')
            continue
        films = {str(f.get('id')): f for f in body.get('films', []) if isinstance(f, dict)}
        events = [e for e in body.get('events', []) if isinstance(e, dict)]
        added = 0
        for event in events:
            film = films.get(str(event.get('filmId')), {})
            item = _quickbook_event_to_web(event, film, day)
            if item['film'] and item['time']:
                web.append(item)
                added += 1
        diagnostics.append(f'{day}: Quickbook {len(events)} eventov, použiteľných {added}')
        if events and i == 0:
            # Compact schema diagnostic helps us finish attribute/version mapping without DevTools.
            ekeys = ', '.join(sorted(events[0].keys()))
            fkeys = ', '.join(sorted(next(iter(films.values())).keys())) if films else '—'
            diagnostics.append(f'Polia eventu: {ekeys}')
            diagnostics.append(f'Polia filmu: {fkeys}')

    # De-duplicate conservatively by the visible screening identity.
    unique, seen = [], set()
    for item in web:
        key = (item['date'], item['time'], norm(item['film']), norm(item['hall']), norm(item['attribute']), norm(item['version']))
        if key in seen:
            continue
        seen.add(key)
        unique.append(item)
    diagnostics.append(f'Quickbook spolu: {len(unique)} predstavení pre {start.isoformat()} – {end.isoformat()}')
    return unique, urls, diagnostics


def scrape_day_safe(cinema, day, movie_titles):
    start = datetime.strptime(day, '%Y-%m-%d').date()
    web, urls, diagnostics = fetch_quickbook_week(cinema, start)
    one = [x for x in web if x['date'] == day]
    return one, (urls[0] if urls else ''), '; '.join(diagnostics)


def scrape_week(cinema, start, movie_titles):
    return fetch_quickbook_week(cinema, start)

def _language_signature(value):
    """Stable signature of Quickbook's structured languages field."""
    try:
        import json
        return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'))
    except Exception:
        return str(value or '')


def _version_key(item):
    """Cinema City screening-language identity, without film/date/time."""
    return str(item.get('_presentation_code') or '').strip().lower()


def learn_quickbook_versions(expected, web):
    """Compatibility hook. V13 reads SUB/DUB directly from Quickbook attributeIds."""
    return []


def same_title(expected, web):
    candidates = [expected.get('film', ''), expected.get('original', '')]
    wn = norm(web.get('film', ''))
    return any(norm(c) == wn for c in candidates if c)


def compatible(e, w):
    return e['date'] == w['date'] and e['time'] == w['time'] and same_title(e, w)


def hall_signature(value):
    """Normalize Excel `Sala 3 VIP` and Quickbook `VIP3` to the same identity."""
    n = norm(value)
    nums = re.findall(r'\d+', n)
    number = nums[-1] if nums else ''
    vip = 'vip' in n
    # A plain number is enough for normal halls; VIP must stay separate.
    return ('vip' if vip else 'hall', number)


def hall_match(expected, actual):
    e = hall_signature(expected)
    w = hall_signature(actual)
    if not e[1] or not w[1]:
        return None
    return e == w


def match_score(e, w):
    """Rank duplicate film+time candidates using hall and screening format.

    Cinema City can have the same film at exactly the same time in VIP and a
    normal/laser hall. V11 paired those rows by list order, creating two false
    errors. The title/date/time remains mandatory; hall and attributes decide
    which duplicate belongs together.
    """
    if not compatible(e, w):
        return None
    score = 0
    hm = hall_match(e.get('hall', ''), w.get('hall', ''))
    if hm is True:
        score += 100
    elif hm is False:
        score -= 100
    ea = set(canon_attributes(e.get('attribute', '')))
    wa = set(canon_attributes(w.get('attribute', '')))
    if ea and wa:
        score += 20 * len(ea & wa)
        score -= 10 * len(ea - wa)
    elif not ea:
        score += 1
    return score


def best_match_index(e, web, used):
    candidates = []
    for i, w in enumerate(web):
        if i in used:
            continue
        score = match_score(e, w)
        if score is not None:
            candidates.append((score, i))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return candidates[0][1]


LASER_BARCO_HALLS = {'5', '6', '7', '8', '9', '10', '12', '13'}
VIP_HALLS = {'1', '2', '3', '4'}

def _hall_number(value):
    sig = hall_signature(value)
    return sig[1] if sig and sig[1] else ''

def attributes_match(expected, actual, hall=''):
    """Compare screening attributes in both directions with Eurovea hall rules.

    Excel `Infinity` and web `Infinity Vision` canonicalize to the same value.
    Laser Barco is special: it may be omitted from Excel, but on the web it is
    required in halls 5,6,7,8,9,10,12,13 and forbidden in every other hall.
    VIP is required on the web in halls 1,2,3,4. Other attributes must agree
    between Excel and Cinema City in both directions.
    """
    e = set(canon_attributes(expected))
    w = set(canon_attributes(actual))
    hn = _hall_number(hall)

    # Hall-specific web requirements.
    if hn in LASER_BARCO_HALLS:
        if 'Laser Barco' not in w:
            return False
    elif 'Laser Barco' in w:
        return False

    if hn in VIP_HALLS and 'VIP' not in w and 'VIP 3D' not in w:
        return False

    # Laser Barco does not have to be written in Excel, so it is excluded from
    # the ordinary Excel<->web equality check. The web rule above is authoritative.
    e.discard('Laser Barco')
    w.discard('Laser Barco')
    return e == w


def _version_parts(value, excel=False):
    """Extract only explicit SUB/DUB and language information.

    For Excel, bare SUB/DUB means Slovak SUB/DUB. Cinema City values stay literal: missing mode or language stays unknown.
    """
    n = norm(value)
    tokens = set(n.split())
    mode = ''
    if {'sub', 'subtitle', 'subtitles'} & tokens or 'titulky' in n:
        mode = 'SUB'
    elif {'dub', 'dubbed', 'dubbing'} & tokens or 'dabing' in n:
        mode = 'DUB'

    lang = ''
    lang_map = {
        'SK': ('sk', 'svk', 'slovak', 'slovencina'),
        'CZ': ('cz', 'cs', 'cze', 'ces', 'czech', 'cestina'),
        'EN': ('en', 'eng', 'english', 'anglictina'),
        'HU': ('hu', 'hun', 'hungarian', 'madarcina'),
        'DE': ('de', 'ger', 'german', 'nemcina'),
    }
    for canonical, variants in lang_map.items():
        if any(norm(v) in tokens for v in variants):
            lang = canonical
            break
    # In the Cinema City SK schedule Excel, a bare SUB/DUB means Slovak.
    # Foreign languages are written explicitly (e.g. Czech SUB).
    if excel and mode and not lang:
        lang = 'SK'
    return mode, lang


def version_match(expected, actual):
    """Compare Excel G with Cinema City localisation exactly.

    Excel bare SUB/DUB means Slovak. A web screening with no SUB/DUB is not the
    same as Excel SUB/DUB; original language is checked separately.
    """
    e_mode, e_lang = _version_parts(expected, excel=True)
    w_mode, w_lang = _version_parts(actual, excel=False)
    if not e_mode and not w_mode:
        return True
    if e_mode != w_mode:
        return False
    if e_mode and e_lang != w_lang:
        return False
    return True


def structured_localization_match(expected_item, web_item):
    """Compare Excel G to Quickbook structured localisation; no reparsing display strings."""
    em = str(expected_item.get('localization_mode') or '').upper()
    el = str(expected_item.get('localization_language') or '').upper()
    wm = str(web_item.get('localization_mode') or '').upper()
    wl = str(web_item.get('localization_language') or '').upper()
    return em == wm and el == wl


def structured_original_language_match(expected_item, web_item):
    """Compare Excel F to Quickbook original language when Excel specifies it."""
    e = str(expected_item.get('original_language_code') or _language_code(expected_item.get('original_language','')) or '').upper()
    w = str(web_item.get('original_language_code') or _language_code(web_item.get('original_language','')) or '').upper()
    if not e:
        return True
    # If Quickbook genuinely omits original language, do not fabricate a different language.
    if not w:
        return True
    return e == w


def explicit_version_conflict(expected, actual):
    """Single source of truth used by all check endpoints."""
    return not version_match(expected, actual)

def filter_noncontradictory_version_errors(errors):
    """V24: version_match is authoritative; do not weaken its result later."""
    return errors


def basic_unmatched(expected, web):
    """Find only film/date/time differences, before attribute/version checks."""
    used = set()
    missing = []
    for e in expected:
        idx = best_match_index(e, web, used)
        if idx is None:
            missing.append(e)
        else:
            used.add(idx)
    extra = [w for i, w in enumerate(web) if i not in used]
    return missing, extra


def reconcile_time_errors(errors):
    """Merge missing+extra rows for the same film/day into one `time` error.

    Hall and format are used to choose the right pair when the same title has
    multiple screenings. This makes a real 12:20 vs 12:00 problem say
    "Nesedí čas" instead of two misleading missing/extra messages.
    """
    missing = [(i, x) for i, x in enumerate(errors) if x.get('type') == 'missing']
    extra = [(i, x) for i, x in enumerate(errors) if x.get('type') == 'extra']
    paired_m, paired_x, merged = set(), set(), []
    for mi, m in missing:
        e = m.get('expected', {})
        candidates = []
        for xi, x in extra:
            if xi in paired_x:
                continue
            w = x.get('web', {})
            if e.get('date') != w.get('date') or not same_title(e, w):
                continue
            score = 0
            hm = hall_match(e.get('hall',''), w.get('hall',''))
            if hm is True: score += 100
            elif hm is False: score -= 100
            ea, wa = set(canon_attributes(e.get('attribute',''))), set(canon_attributes(w.get('attribute','')))
            score += 20 * len(ea & wa)
            score -= 10 * len(ea - wa)
            # Prefer the nearest time if metadata ties.
            try:
                eh, em = map(int, e.get('time','0:0').split(':'))
                wh, wm = map(int, w.get('time','0:0').split(':'))
                distance = abs((eh*60+em)-(wh*60+wm))
            except Exception:
                distance = 9999
            candidates.append((score, -distance, xi, x))
        if candidates:
            candidates.sort(reverse=True, key=lambda z:(z[0],z[1]))
            score, _, xi, x = candidates[0]
            # Require either matching hall/format evidence, or a unique same-title candidate.
            if score > 0 or len(candidates) == 1:
                paired_m.add(mi); paired_x.add(xi)
                merged.append({'type':'time','expected':e,'web':x.get('web',{})})
    out=[]
    for i,x in enumerate(errors):
        if (x.get('type')=='missing' and i in paired_m) or (x.get('type')=='extra' and i in paired_x):
            continue
        out.append(x)
    out.extend(merged)
    return out


def reconcile_title_errors(errors):
    """Merge missing+extra rows at the same date/time/hall into one title error."""
    missing = [(i, x) for i, x in enumerate(errors) if x.get('type') == 'missing']
    extra = [(i, x) for i, x in enumerate(errors) if x.get('type') == 'extra']
    paired_m, paired_x, merged = set(), set(), []
    for mi, m in missing:
        e = m.get('expected', {})
        candidates = []
        for xi, x in extra:
            if xi in paired_x:
                continue
            w = x.get('web', {})
            if e.get('date') != w.get('date') or e.get('time') != w.get('time'):
                continue
            hm = hall_match(e.get('hall',''), w.get('hall',''))
            if hm is True:
                candidates.append((xi, x))
        if len(candidates) == 1:
            xi, x = candidates[0]
            paired_m.add(mi); paired_x.add(xi)
            merged.append({'type':'title','expected':e,'web':x.get('web',{})})
    out=[]
    for i,x in enumerate(errors):
        if (x.get('type')=='missing' and i in paired_m) or (x.get('type')=='extra' and i in paired_x):
            continue
        out.append(x)
    out.extend(merged)
    return out


def compact_show(item):
    return {
        'date': item.get('date', ''), 'time': item.get('time', ''),
        'film': item.get('film', ''), 'hall': item.get('hall', ''),
        'attribute': item.get('attribute', ''), 'original_language': item.get('original_language', ''),
        'version': item.get('version', ''),
        'original_language_code': item.get('original_language_code', ''),
        'localization_mode': item.get('localization_mode', ''),
        'localization_language': item.get('localization_language', ''),
    }



@app.after_request
def no_cache(response):
    # During active development never let Render/browser serve an older API/static response.
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

@app.get('/api/version')
def api_version():
    cases = [
        ('SUB', 'Slovak SUB', True),
        ('SUB', 'Czech SUB', False),
        ('Czech SUB', 'Czech SUB', True),
        ('Czech SUB', 'Slovak SUB', False),
        ('Czech DUB', 'DUB CS', True),
        ('SUB', 'Slovak', True),
    ]
    tests = [{'excel': a, 'web': b, 'expected': want, 'got': version_match(a,b),
              'ok': version_match(a,b) == want} for a,b,want in cases]
    return jsonify({'ok': all(x['ok'] for x in tests), 'build': BUILD_VERSION, 'tests': tests})

@app.get('/')
def home():
    return render_template('index.html')


@app.get('/health')
def health():
    return {'ok': True}


def _selection_key(item):
    """Stable film-version key used only by the V30 UI filter."""
    title = norm(item.get('film', ''))
    original = item.get('original_language_code') or _language_code(item.get('original_language', '')) or ''
    mode = item.get('localization_mode') or _version_parts(item.get('version', ''), excel=True)[0] or ''
    lang = item.get('localization_language') or _version_parts(item.get('version', ''), excel=True)[1] or ''
    attrs = '|'.join(sorted(canon_attributes(item.get('attribute', ''))))
    return '::'.join((title, original, mode, lang, attrs))


def _selected_keys_from_request():
    raw = request.form.get('selected_variants', '').strip()
    if not raw:
        return None
    return {x for x in raw.split('|||') if x}


def _selected_halls_from_request():
    raw = request.form.get('selected_halls', '').strip()
    if not raw:
        return None
    return {x for x in raw.split('|||') if x}


def _hall_number(value):
    nums = re.findall(r'\d+', norm(value))
    return nums[-1] if nums else ''


def _reserve_unselected_web(unselected_expected, web):
    """Reserve web rows belonging to film versions the user did not select.

    This prevents another SUB/DUB/format of the same title from appearing as an
    'extra' error while a single variant is being checked.
    """
    reserved = set()
    for e in unselected_expected:
        idx = best_match_index(e, web, reserved)
        if idx is not None:
            reserved.add(idx)
    return reserved


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
        for x in exp:
            x['selection_key'] = _selection_key(x)
        return jsonify({'ok': True, 'expected': exp, 'count': len(exp)})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 400


@app.post('/api/web-preview-day')
def web_preview_day():
    """V7 compatibility diagnostic endpoint.

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
        day_exp = [x for x in exp if x['date'] == day]
        learn_quickbook_versions(day_exp, web)
        expected_count = len(day_exp)
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
        version_notes = learn_quickbook_versions(exp, web)
        diagnostics.extend(version_notes[:20])
        by_day, expected_by_day = {}, {}
        for x in web:
            by_day[x['date']] = by_day.get(x['date'], 0) + 1
        for x in exp:
            expected_by_day[x['date']] = expected_by_day.get(x['date'], 0) + 1
        missing_basic, extra_basic = basic_unmatched(exp, web)
        diagnostics.append(f'Film+dátum+čas: v Exceli navyše {len(missing_basic)}, na Cinema City navyše {len(extra_basic)}')
        for item in missing_basic[:12]:
            diagnostics.append(f"EXCEL NAVYŠE: {item['date']} {item['time']} | {item['film']} | {item.get('hall','')}")
        for item in extra_basic[:12]:
            diagnostics.append(f"WEB NAVYŠE: {item['date']} {item['time']} | {item['film']} | {item.get('hall','')}")
        return jsonify({
            'ok': True, 'web': web, 'count': len(web), 'by_day': by_day,
            'expected_by_day': expected_by_day,
            'missing_basic': [compact_show(x) for x in missing_basic],
            'extra_basic': [compact_show(x) for x in extra_basic],
            'diagnostics': diagnostics, 'urls': urls,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@app.post('/api/check-day')
def check_day():
    """Check one visible day so the browser can show truthful Thu→Wed progress."""
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
        day_index = int(request.form.get('day_index', '0'))
        if day_index < 0 or day_index > 6:
            raise ValueError('Neplatný deň kontroly.')
        exp_all = expected_from_excel(f.read(), cinema, start)
        day = (start + timedelta(days=day_index)).isoformat()
        day_all = [x for x in exp_all if x['date'] == day]
        selected_keys = _selected_keys_from_request()
        selected_halls = _selected_halls_from_request()
        hall_scope = day_all if selected_halls is None else [x for x in day_all if _hall_number(x.get('hall','')) in selected_halls]
        exp = hall_scope if selected_keys is None else [x for x in hall_scope if _selection_key(x) in selected_keys]
        unselected = [] if selected_keys is None else [x for x in hall_scope if _selection_key(x) not in selected_keys]
        titles = known_titles_from_expected(exp_all)
        web_all, url, diagnostic = scrape_day_safe(cinema, day, titles)
        web = web_all if selected_halls is None else [x for x in web_all if _hall_number(x.get('hall','')) in selected_halls]
        notes = learn_quickbook_versions(exp, web)
        if not web_all and exp:
            return jsonify({'ok': False, 'error': f'Cinema City nevrátilo dáta pre {day}.', 'diagnostic': diagnostic}), 502
        errors = []
        used = _reserve_unselected_web(unselected, web)
        for e in exp:
            idx = best_match_index(e, web, used)
            if idx is None:
                errors.append({'type': 'missing', 'expected': compact_show(e)})
                continue
            used.add(idx); w = web[idx]
            hm = hall_match(e.get('hall', ''), w.get('hall', ''))
            if hm is False:
                errors.append({'type': 'hall', 'expected': compact_show(e), 'web': compact_show(w)})
            if not attributes_match(e['attribute'], w['attribute'], e.get('hall', '') or w.get('hall', '')):
                errors.append({'type': 'attribute', 'expected': compact_show(e), 'web': compact_show(w)})
            if not structured_original_language_match(e, w):
                errors.append({'type': 'original_language', 'expected': compact_show(e), 'web': compact_show(w)})
            if not version_match(e.get('version',''), w.get('version','')):
                errors.append({'type': 'version', 'expected': compact_show(e), 'web': compact_show(w)})
        selected_title_norms = {norm(x.get('film','')) for x in exp}
        for i, w in enumerate(web):
            same_slot_as_selected = any(
                e.get('date') == w.get('date') and e.get('time') == w.get('time')
                and hall_match(e.get('hall',''), w.get('hall','')) is True
                for e in exp
            )
            if i not in used and (selected_keys is None or norm(w.get('film','')) in selected_title_norms or same_slot_as_selected):
                errors.append({'type': 'extra', 'web': compact_show(w)})
        errors = reconcile_time_errors(errors)
        errors = reconcile_title_errors(errors)
        errors = filter_noncontradictory_version_errors(errors)
        return jsonify({'ok': True, 'day': day, 'expected': len(exp), 'web': len(web),
                        'errors': errors, 'diagnostic': diagnostic, 'version_notes': notes, 'url': url, 'build': BUILD_VERSION})
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
        version_notes = learn_quickbook_versions(exp, web)
        diagnostics.extend(version_notes[:20])
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
            idx = best_match_index(e, web, used)
            if idx is None:
                errors.append({'type': 'missing', 'expected': e})
                continue
            used.add(idx)
            w = web[idx]
            if not attributes_match(e['attribute'], w['attribute'], e.get('hall', '') or w.get('hall', '')):
                errors.append({'type': 'attribute', 'expected': e, 'web': w})
            if not structured_original_language_match(e, w):
                errors.append({'type': 'original_language', 'expected': e, 'web': w})
            if not version_match(e.get('version',''), w.get('version','')):
                errors.append({'type': 'version', 'expected': e, 'web': w})

        for i, w in enumerate(web):
            if i not in used:
                errors.append({'type': 'extra', 'web': w})
        errors = reconcile_time_errors(errors)
        errors = reconcile_title_errors(errors)
        errors = filter_noncontradictory_version_errors(errors)

        return jsonify({
            'ok': True, 'expected': len(exp), 'web': len(web), 'errors': errors,
            'urls': urls, 'diagnostics': diagnostics,
        })
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', '5000')), debug=True)
