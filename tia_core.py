# -*- coding: utf-8 -*-
"""Traduttore TIA - nucleo: lettura/scrittura Excel TIA, memoria di traduzione,
glossario, controlli, motore LM Studio (API compatibile OpenAI)."""
import os, sys, re, json, csv, zipfile, time, random, urllib.request, urllib.error
from xml.sax.saxutils import escape, unescape

# da exe (PyInstaller) i file dati stanno accanto all'eseguibile
BASE = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
LANGNAMES = {'en-US': 'inglese (en-US)', 'de-DE': 'tedesco (de-DE)', 'fr-FR': 'francese (fr-FR)',
             'es-ES': 'spagnolo (es-ES)', 'pt-PT': 'portoghese (pt-PT)', 'nl-NL': 'olandese (nl-NL)',
             'pl-PL': 'polacco (pl-PL)', 'cs-CZ': 'ceco (cs-CZ)', 'ro-RO': 'rumeno (ro-RO)',
             'tr-TR': 'turco (tr-TR)', 'ru-RU': 'russo (ru-RU)', 'zh-CN': 'cinese (zh-CN)'}
LANG_RE = re.compile(r'^[a-z]{2,3}-[A-Z]{2}\*?$')
FIELD_RE = re.compile(r'<field[^>]*/>')

def has_letters(s):
    return re.search(r'[A-Za-zÀ-ÿ]{3,}', FIELD_RE.sub('', s)) is not None

def is_placeholder(cur, src):
    """True se la cella target e' vuota o contiene un segnaposto TIA da sostituire."""
    if cur is None or cur.strip() == '':
        return True
    if cur == src:
        return True
    if cur.strip() == 'Text' and src.strip() != 'Text':
        return True
    only_fields = not re.sub(r'<field[^>]*/>|[\s:;,.\-]', '', cur)
    src_only_fields = not re.sub(r'<field[^>]*/>|[\s:;,.\-]', '', src)
    return only_fields and not src_only_fields

def split_ws(s):
    m = re.match(r'^(\s*)(.*?)(\s*)$', s, re.S)
    return m.group(1), m.group(2), m.group(3)

# ---------------------------------------------------------------- configurazione
DEFAULT_CONFIG = {
    "url": "http://localhost:1234/v1/chat/completions",
    "modello": "openai/gpt-oss-120b",
    "dimensione_lotto": 10,
    "temperatura": 0.2,
    "reasoning_effort": "low",
    "timeout_s": 1800,
    "lingua_origine": "it-IT",
    "percentuale_campione": 10,
    "applica_bozze_da_rivedere": False,
    "esempi_nel_prompt": 12,
    "file_tm": "memoria_traduzione.json",
    "file_glossario": "glossario.txt",
    "file_classi": "classi_allarme.json",
    "categorie_allarme": ["Alarm", "ProDiag", "testi di segnalazione", "text list"],
}

def load_config(path=None):
    path = path or os.path.join(BASE, 'config.json')
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        try:
            cfg.update(json.load(open(path, encoding='utf-8')))
        except Exception:
            pass
    return cfg

def save_config(cfg, path=None):
    path = path or os.path.join(BASE, 'config.json')
    json.dump(cfg, open(path, 'w', encoding='utf-8'), ensure_ascii=False, indent=2)

def rel(p):
    return p if os.path.isabs(p) else os.path.join(BASE, p)

# ---------------------------------------------------------------- memoria di traduzione
class TM:
    """entries[testo_origine][lingua] = {"t": traduzione, "s": stato}
    stato: 'rivisto' (corretto/approvato da te) > 'modello' > 'excel' (appreso da progetti)"""
    PRIO = {'rivisto': 3, 'modello': 2, 'excel': 1}

    def __init__(self, path):
        self.path = path
        self.entries = {}
        if os.path.exists(path):
            self.entries = json.load(open(path, encoding='utf-8')).get('entries', {})

    def save(self):
        json.dump({'versione': 1, 'entries': self.entries}, open(self.path, 'w', encoding='utf-8'),
                  ensure_ascii=False, indent=0)

    def get(self, core, lang):
        e = self.entries.get(core, {}).get(lang)
        return e['t'] if e else None

    def put(self, core, lang, text, status):
        cur = self.entries.setdefault(core, {}).get(lang)
        if cur and self.PRIO.get(cur['s'], 0) > self.PRIO.get(status, 0):
            return False
        self.entries[core][lang] = {'t': text, 's': status}
        return True

    def similar(self, core, lang, n):
        """Voci della TM piu' simili (parole in comune): usate come esempi nel prompt."""
        if n <= 0:
            return []
        w = set(re.findall(r'\w+', core.lower()))
        if not w:
            return []
        scored = []
        for k, v in self.entries.items():
            if lang not in v or k == core:
                continue
            kw = set(re.findall(r'\w+', k.lower()))
            if not kw:
                continue
            j = len(w & kw) / len(w | kw)
            if j > 0.2:
                scored.append((j, k))
        scored.sort(reverse=True)
        return [(k, self.entries[k][lang]['t']) for _, k in scored[:n]]

# ---------------------------------------------------------------- glossario
class Glossary:
    """Righe: termine_it | trad_lingua1 | trad_lingua2 ...   (prima riga '# lingue: en-US, de-DE')
    Alternative con '/'. '=' = invariato."""
    def __init__(self, path):
        self.path = path
        self.langs = []
        self.terms = []   # (alternative_it, {lang: [alternative]})
        self.text = ''
        if os.path.exists(path):
            self.text = open(path, encoding='utf-8').read()
            self._parse()

    def _parse(self):
        for ln in self.text.splitlines():
            ln = ln.strip()
            if not ln:
                continue
            if ln.lower().startswith('# lingue:'):
                self.langs = [x.strip() for x in ln.split(':', 1)[1].split(',')]
                continue
            if ln.startswith('#'):
                continue
            parts = [p.strip() for p in ln.split('|')]
            if len(parts) < 2:
                continue
            its = [x.strip() for x in parts[0].split('/') if x.strip()]
            tr = {}
            for lang, p in zip(self.langs, parts[1:]):
                tr[lang] = [x.strip() for x in p.split('/') if x.strip()]
            self.terms.append((its, tr))

    def for_prompt(self, langs):
        lines = []
        for its, tr in self.terms:
            seg = []
            for l in langs:
                if l in tr:
                    seg.append(' / '.join(tr[l]))
            lines.append('%s = %s' % (' / '.join(its), ' ; '.join(seg)))
        return '\n'.join(lines)

    def check(self, src, tr, lang):
        out = []
        low = tr.lower()
        for its, trs in self.terms:
            if lang not in trs:
                continue
            for it in its:
                if re.search(r'(?<!\w)' + re.escape(it) + r'(?!\w)', src, re.I):
                    exp = [('' if e == '=' else e) or it for e in trs[lang]]
                    exp = [it if e == '=' else e for e in trs[lang]]
                    if not any(e.lower() in low for e in exp):
                        out.append('glossario: "%s" -> %s' % (it, ' / '.join(exp)))
                    break
        return out

# ---------------------------------------------------------------- controlli
IT_WORDS = {'della', 'delle', 'degli', 'dello', 'nella', 'nelle', 'sono', 'viene', 'dopo', 'prima',
            'posizione', 'allarme', 'allarmi', 'velocità', 'tubo', 'attesa', 'comando', 'ciclo',
            'abilitato', 'raggiunta', 'raggiunto', 'dalla', 'dallo', 'quando', 'senza', 'anche'}

def prefix_tokens(s):
    m = re.match(r'^((?:SR|TR|S|T|sr|tr|s|t)\s+)', s)
    lead = re.match(r'^([<>*!\-0-9\s]*)', s).group(1) if s else ''
    return (m.group(1).strip() if m else ''), lead.strip()

def check_translation(src, tr, lang, glossary):
    """Ritorna la lista dei motivi per cui la traduzione va rivista (vuota = ok)."""
    r = []
    if not tr.strip():
        return ['vuota']
    if FIELD_RE.findall(src) != FIELD_RE.findall(tr):
        r.append('segnaposto <field> diversi')
    if src.count('\n') != tr.count('\n'):
        r.append('numero di righe diverso')
    p_s, l_s = prefix_tokens(src)
    p_t, l_t = prefix_tokens(tr)
    if p_s and p_s != p_t:
        r.append('prefisso "%s" non mantenuto' % p_s)
    if l_s and l_s != l_t and not p_s:
        r.append('simboli iniziali "%s" non mantenuti' % l_s)
    words = set(re.findall(r'[A-Za-zÀ-ÿ]+', tr.lower()))
    left = sorted(words & IT_WORDS)
    if left and not lang.startswith('it'):
        r.append('parole italiane rimaste: ' + ', '.join(left))
    if tr.strip() == src.strip() and has_letters(src) and len(src) > 4:
        r.append('identica all\'originale')
    if len(tr) > max(4 * len(src), 40) and len(tr) > 3 * len(src):
        r.append('molto piu\' lunga dell\'originale')
    r += glossary.check(src, tr, lang)
    return r

# ---------------------------------------------------------------- regole "non tradurre"
def is_copy_rule(core):
    """Testi che si copiano invariati (codici, identificatori, markup di sistema)."""
    if not has_letters(core):
        return True
    if '@' in core and '%' in core:      # testi ProDiag @1%t#...@
        return True
    if '$$' in core:
        return True
    if re.fullmatch(r'[A-Za-z0-9_.\[\]()\-:,/+ ]{1,40}', core) and ' ' not in core.strip() and \
            re.search(r'[0-9_.\[\]]', core):
        return True
    return False

# ---------------------------------------------------------------- progetto Excel
class Project:
    def __init__(self, path, cfg):
        import openpyxl
        self.path = path
        self.wb = openpyxl.load_workbook(path)
        self.ws = self.wb['User Texts'] if 'User Texts' in self.wb.sheetnames else self.wb.worksheets[0]
        hdr = [c.value for c in self.ws[1]]
        self.hdr = hdr
        self.cols = {}   # lingua -> lettera colonna
        self.ref_col = None
        for i, h in enumerate(hdr):
            if h and LANG_RE.match(str(h)):
                letter = openpyxl.utils.get_column_letter(i + 1)
                if str(h).endswith('*'):
                    self.ref_col = letter
                else:
                    self.cols[str(h)] = letter
        self.src_lang = cfg.get('lingua_origine', 'it-IT')
        if self.src_lang not in self.cols:
            raise ValueError('Colonna lingua origine "%s" non trovata. Colonne: %s' % (self.src_lang, list(self.cols)))
        self.targets = [l for l in self.cols if l != self.src_lang]
        self.rows = []
        idx = {l: hdr.index(l) for l in self.cols}
        for n, r in enumerate(self.ws.iter_rows(min_row=2, values_only=True), start=2):
            src = r[idx[self.src_lang]] or ''
            if not str(src).strip():
                continue
            self.rows.append({'row': n, 'cat': r[0] or '', 'vp': r[1] or '', 'src': str(src),
                              'cur': {l: (r[idx[l]] or '') for l in self.targets}})

# ---------------------------------------------------------------- classi di allarme
def load_classes(path):
    if os.path.exists(path):
        return json.load(open(path, encoding='utf-8'))
    return {}

def class_row_info(vp, cat):
    if 'AlarmClassDisplayName' not in cat:
        return None
    parts = vp.split('\\')
    if parts[-1] == 'ShortName':
        return parts[-2], 'sigla'
    if parts[-1].startswith('AlarmClassData_IDisplayNaming'):
        return parts[-2], 'nome'
    return None

# ---------------------------------------------------------------- piano di traduzione
def build_plan(project, tm, classes, targets, only_cats=None):
    """Calcola le celle da scrivere usando regole+classi+TM; elenca i testi ancora da tradurre."""
    edits = {}      # (riga, lettera) -> (testo, origine)
    pending = {}    # core -> set(lingue mancanti)
    notes = []
    stats = {'righe': len(project.rows), 'celle_da_compilare': 0, 'tm': 0, 'regole': 0, 'classi': 0}
    for r in project.rows:
        lead, core, trail = split_ws(r['src'])
        ci = class_row_info(r['vp'], r['cat'])
        for lang in targets:
            cur = r['cur'].get(lang, '')
            if ci:
                # nelle righe delle classi il segnaposto e' il nome originale della classe
                if not (is_placeholder(cur, r['src']) or cur.strip() == ci[0]):
                    continue
            elif not is_placeholder(cur, r['src']):
                continue
            stats['celle_da_compilare'] += 1
            letter = project.cols[lang]
            if ci:
                name, kind = ci
                val = classes.get(name, {}).get(lang, {}).get(kind)
                if val is None:
                    if cur.strip() == '' or cur == r['src'] or cur.strip() == name:
                        notes.append('classe di allarme "%s" non configurata (%s)' % (name, lang))
                    continue
                if cur != val:
                    edits[(r['row'], letter)] = (val, 'classi'); stats['classi'] += 1
                continue
            if is_copy_rule(core):
                if cur != r['src']:
                    edits[(r['row'], letter)] = (r['src'], 'regole'); stats['regole'] += 1
                continue
            t = tm.get(core, lang)
            if t is not None:
                if t == '=':
                    t = core
                val = lead + t + trail
                if cur != val:
                    edits[(r['row'], letter)] = (val, 'tm'); stats['tm'] += 1
            else:
                if only_cats and not any(c in r['cat'] or c in r['vp'] for c in only_cats):
                    stats['fuori_filtro'] = stats.get('fuori_filtro', 0) + 1
                    continue
                pending.setdefault(core, set()).add(lang)
    return edits, pending, stats, sorted(set(notes))

# ---------------------------------------------------------------- motore (LM Studio / OpenAI-compatibile)
SYS_PROMPT = """Sei un traduttore tecnico di automazione industriale (Siemens TIA Portal, macchine utensili).
Traduci dall'italiano nelle lingue indicate i testi numerati.
Regole:
- Usa il glossario. Stile conciso da testo di macchina/allarme, come negli esempi (maiuscole come nell'originale, non una per parola).
- Non tradurre sigle, codici, nomi di variabili, numeri, ne' i prefissi S / T / TR / SR davanti al testo.
- Conserva simboli iniziali (>, <, *, 0>1, 1>0 ...), la punteggiatura e i segnaposto <field ref="n" /> invariati.
- Se il testo e' gia' nella lingua richiesta o e' un identificatore, ripetilo invariato.
- Nessuna spiegazione. Rispondi SOLO con righe nel formato:
N ## traduzione_1 ## traduzione_2 ...
(nell'ordine delle lingue indicate, un a capo si scrive \\n)
"""

def chat(cfg, prompt):
    body = {'model': cfg['modello'], 'messages': [{'role': 'user', 'content': prompt}],
            'temperature': cfg.get('temperatura', 0.2)}
    if cfg.get('reasoning_effort'):
        body['reasoning_effort'] = cfg['reasoning_effort']
    req = urllib.request.Request(cfg['url'], data=json.dumps(body).encode('utf-8'),
                                 headers={'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=cfg.get('timeout_s', 1800)) as f:
        d = json.loads(f.read().decode('utf-8'))
    return d['choices'][0]['message']['content'], d.get('usage', {}).get('completion_tokens', 0)

def clean(s):
    s = s.replace('‑', '-').replace('‐', '-').replace(' ', ' ')
    return s.strip()

def translate_batch(cfg, cores, langs, tm, glossary):
    """cores: lista di testi origine (puliti). I testi multiriga sono tradotti riga per riga."""
    segs = []   # (indice_core, indice_riga, testo)
    for i, c in enumerate(cores):
        for j, line in enumerate(c.split('\n')):
            segs.append((i, j, line))
    items = [(k + 1, s) for k, s in enumerate(segs) if s[2].strip()]
    ex = []
    n_ex = cfg.get('esempi_nel_prompt', 12)
    seen = set()
    for c in cores:
        for k, v in tm.similar(c.split('\n')[0], langs[0], max(2, n_ex // max(1, len(cores)) + 1)):
            if k not in seen and len(ex) < n_ex and '\n' not in k:
                seen.add(k)
                ex.append('%s ## %s' % (k, ' ## '.join((tm.get(k, l) or '').replace('\n', '\\n') for l in langs)))
    prompt = SYS_PROMPT + '\nLingue (in quest\'ordine): ' + ' ; '.join(LANGNAMES.get(l, l) for l in langs) + \
        '\n\nGlossario:\n' + glossary.for_prompt(langs)
    if ex:
        prompt += '\n\nEsempi gia\' approvati (italiano ## ' + ' ## '.join(langs) + '):\n' + '\n'.join(ex)
    prompt += '\n\nTESTI:\n' + '\n'.join('%d ## %s' % (n, s[2].replace('\n', '\\n')) for n, s in items) + '\n'
    out, ntok = chat(cfg, prompt)
    parsed = {}
    for ln in out.splitlines():
        m = re.match(r'\s*(\d+)\s*##\s*(.*)$', ln)
        if not m:
            continue
        parts = [clean(p).replace('\\n', '\n') for p in m.group(2).split(' ## ')]
        if len(parts) == len(langs):
            parsed[int(m.group(1))] = parts
    res = [{l: [''] * len(c.split('\n')) for l in langs} for c in cores]
    bad = set()
    for n, (i, j, line) in items:
        if n in parsed:
            for l, t in zip(langs, parsed[n]):
                res[i][l][j] = t
        else:
            bad.add(i)
    for i, c in enumerate(cores):
        for j, line in enumerate(c.split('\n')):
            if not line.strip():   # righe vuote conservate
                for l in langs:
                    res[i][l][j] = ''
    out_list = []
    for i, c in enumerate(cores):
        out_list.append(({l: '\n'.join(res[i][l]) for l in langs}, i in bad))
    return out_list, ntok

# ---------------------------------------------------------------- residuo (revisione)
RES_HEAD = ['id', 'italiano']

def export_residual(path, items, langs):
    """items: lista di dict {core, bozza{lang}, motivi[]}"""
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['id', 'italiano'] + ['bozza ' + l for l in langs] + ['motivo'])
        for n, it in enumerate(items, 1):
            w.writerow([n, it['core']] + [it['bozza'].get(l, '') for l in langs] + [' | '.join(it['motivi'])])

def import_corrections(path, tm, langs):
    n = 0
    with open(path, newline='', encoding='utf-8-sig') as f:
        rd = csv.reader(f, delimiter=';')
        head = next(rd)
        cols = {}
        for l in langs:
            for i, h in enumerate(head):
                if h.strip() in ('bozza ' + l, l):
                    cols[l] = i
        for row in rd:
            if len(row) < 2 or not row[1].strip():
                continue
            core = row[1].strip() if '\n' not in row[1] else row[1].strip()
            for l, i in cols.items():
                if i < len(row) and row[i].strip():
                    tm.put(core, l, row[i].strip('\r') if '\n' in row[i] else row[i].strip(), 'rivisto')
                    n += 1
    return n

# ---------------------------------------------------------------- apprendimento da Excel gia' tradotto
def learn_from_project(project, tm, targets):
    n = 0
    for r in project.rows:
        core = r['src'].strip()
        if is_copy_rule(core):
            continue
        for l in targets:
            cur = r['cur'].get(l, '').strip()
            if cur and not is_placeholder(r['cur'][l], r['src']) and not re.fullmatch(r'(<field[^>]*/>)+', cur):
                if tm.get(core, l) is None:
                    tm.put(core, l, cur, 'excel'); n += 1
    return n

def consistency_report(project, targets):
    """Stesso testo italiano con traduzioni diverse dentro lo stesso Excel."""
    d = {}
    for r in project.rows:
        core = r['src'].strip()
        for l in targets:
            cur = r['cur'].get(l, '').strip()
            if cur and not is_placeholder(r['cur'][l], r['src']):
                d.setdefault((core, l), {}).setdefault(cur, 0)
                d[(core, l)][cur] += 1
    return [(c, l, v) for (c, l), v in d.items() if len(v) > 1]

# ---------------------------------------------------------------- scrittura Excel (patch XML)
def _sheet_path(z, sheetname):
    wbx = z.read('xl/workbook.xml').decode('utf-8-sig')
    m = re.search(r'<(?:\w+:)?sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheetname), wbx) or \
        re.search(r'<(?:\w+:)?sheet [^>]*r:id="([^"]+)"[^>]*name="%s"' % re.escape(sheetname), wbx)
    rid = m.group(1)
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8-sig')
    t = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    tgt = t.group(1)
    return tgt.lstrip('/') if tgt.startswith('/') else 'xl/' + tgt

def write_edits(src, dst, sheetname, edits):
    """Modifica solo le celle indicate patchando sharedStrings e foglio (resto del file identico)."""
    zin = zipfile.ZipFile(src)
    sp = _sheet_path(zin, sheetname)
    ss = zin.read('xl/sharedStrings.xml').decode('utf-8-sig')
    sh = zin.read(sp).decode('utf-8-sig')
    px = re.search(r'<((?:\w+:)?)sst[ >]', ss).group(1)
    n0 = int(re.search(r'uniqueCount="(\d+)"', ss).group(1))
    assert ss.count('<%ssi>' % px) == n0, 'sharedStrings con formato inatteso'
    new, idx = [], {}
    for t in edits.values():
        if t not in idx:
            idx[t] = n0 + len(new); new.append(t)
    add = ''.join('<%ssi><%st xml:space="preserve">%s</%st></%ssi>' % (
        px, px, escape(t).replace('\r\n', '\n').replace('\n', '\r\n'), px, px) for t in new)
    ss = ss.replace('</%ssst>' % px, add + '</%ssst>' % px)
    ss = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % (n0 + len(new)), ss, count=1)
    ok = 0
    skipped = []
    for (row, col), t in edits.items():
        pat = re.compile(r'(<%sc r="%s%d"(?: s="\d+")? t="s"><%sv>)(\d+)(</%sv></%sc>)' % (px, col, row, px, px, px))
        if len(pat.findall(sh)) != 1:
            skipped.append((row, col)); continue
        sh = pat.sub(lambda m: m.group(1) + str(idx[t]) + m.group(3), sh, count=1); ok += 1
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for i in zin.infolist():
            d = zin.read(i.filename)
            if i.filename == 'xl/sharedStrings.xml':
                d = ('﻿' + ss).encode('utf-8')
            elif i.filename == sp:
                d = ('﻿' + sh).encode('utf-8')
            zo.writestr(i, d)
    return ok, skipped

def verify_output(src, dst, sheetname, edits_cols):
    """Controlla che siano cambiate solo le celle previste. Ritorna (n_modificate, n_anomalie)."""
    import openpyxl
    def sheet(p):
        wb = openpyxl.load_workbook(p)
        return wb[sheetname] if sheetname in wb.sheetnames else wb.worksheets[0]
    a, b = sheet(src), sheet(dst)
    diff = bad = 0
    for ra, rb in zip(a.iter_rows(), b.iter_rows()):
        for ca, cb in zip(ra, rb):
            if ca.value != cb.value:
                diff += 1
                exp = edits_cols.get((ca.row, ca.column_letter))
                if exp is None or exp.replace('\r\n', '\n') != (cb.value or '').replace('\r\n', '\n'):
                    bad += 1
    return diff, bad

# ---------------------------------------------------------------- sessione di lavoro
class Session:
    def __init__(self, cfg=None, log=print):
        self.cfg = cfg or load_config()
        self.log = log
        self.tm = TM(rel(self.cfg['file_tm']))
        self.glossary = Glossary(rel(self.cfg['file_glossario']))
        self.classes = load_classes(rel(self.cfg['file_classi']))
        self.project = None
        self.targets = []
        self.drafts = {}      # core -> {lingua: testo} (bozze non validate)
        self.review = []      # [{core, bozza, motivi}]
        self.stop = False
        self.only_alarms = False

    def reload_resources(self):
        self.tm = TM(rel(self.cfg['file_tm']))
        self.glossary = Glossary(rel(self.cfg['file_glossario']))
        self.classes = load_classes(rel(self.cfg['file_classi']))

    def open(self, path):
        self.project = Project(path, self.cfg)
        self.targets = list(self.project.targets)
        self.drafts, self.review = {}, []
        return self.project

    def plan(self):
        cats = self.cfg.get('categorie_allarme') if self.only_alarms else None
        return build_plan(self.project, self.tm, self.classes, self.targets, cats)

    def translate_pending(self, progress=None):
        """Traduce con il modello i testi non coperti dalla TM. Salva in TM quelli che superano i controlli."""
        edits, pending, stats, notes = self.plan()
        cores = sorted(pending)
        langs = list(self.targets)
        bs = int(self.cfg.get('dimensione_lotto', 10))
        pct = float(self.cfg.get('percentuale_campione', 0)) / 100.0
        self.stop = False
        done = 0
        t0 = time.time()
        tokens = 0
        for i in range(0, len(cores), bs):
            if self.stop:
                self.log('Interrotto dall\'utente.')
                break
            batch = cores[i:i + bs]
            try:
                res, nt = translate_batch(self.cfg, batch, langs, self.tm, self.glossary)
            except Exception as e:
                self.log('ERRORE di connessione al modello: %s' % e)
                raise
            tokens += nt
            for core, (tr, missing) in zip(batch, res):
                motivi = []
                if missing:
                    motivi.append('risposta incompleta')
                for l in langs:
                    if l in pending[core] or True:
                        motivi += ['%s: %s' % (l, m) for m in check_translation(core, tr[l], l, self.glossary)]
                if motivi:
                    self.drafts[core] = tr
                    self.review.append({'core': core, 'bozza': tr, 'motivi': motivi})
                else:
                    for l in pending[core]:
                        self.tm.put(core, l, tr[l], 'modello')
                    if pct and random.random() < pct:
                        self.review.append({'core': core, 'bozza': tr, 'motivi': ['campione di controllo']})
            done += len(batch)
            self.tm.save()   # salvataggio a ogni lotto: nulla va perso se l'app si chiude
            if progress:
                progress(done, len(cores), time.time() - t0, tokens)
        self.tm.save()
        return len(cores), len(self.review)

    def save_excel(self, out_path):
        edits, pending, stats, notes = self.plan()
        final = {k: v[0] for k, v in edits.items()}
        if self.cfg.get('applica_bozze_da_rivedere'):
            for r in self.project.rows:
                core = r['src'].strip()
                if core in self.drafts:
                    lead, _, trail = split_ws(r['src'])
                    for l in self.targets:
                        if is_placeholder(r['cur'][l], r['src']):
                            final[(r['row'], self.project.cols[l])] = lead + self.drafts[core][l] + trail
        sheet = self.project.ws.title
        ok, skipped = write_edits(self.project.path, out_path, sheet, final)
        diff, bad = verify_output(self.project.path, out_path, sheet, final)
        return {'scritte': ok, 'saltate': len(skipped), 'modificate': diff, 'anomalie': bad,
                'ancora_da_tradurre': len([c for c in pending if c not in self.drafts or not self.cfg.get('applica_bozze_da_rivedere')])}
