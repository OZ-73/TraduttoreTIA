# -*- coding: utf-8 -*-
"""Traduttore TIA - controllo delle traduzioni GIA' PRESENTI nell'Excel e applicazione delle correzioni.

Il controllo e' strutturale e lessicale (niente modello): segnaposto, righe, numeri, parole italiane rimaste, testi
predefiniti TIA, contraddizioni (max/min, apri/chiudi, avanti/indietro...), glossario, errori noti (errori_noti.txt),
traduzioni mancanti, coerenza interna. Non capisce il senso di una frase: per quello serve la rilettura di Claude in chat
sul CSV prodotto. Le correzioni tornano nel file con la colonna "applica" e si scrivono sull'Excel solo se la cella
contiene ancora il testo controllato."""
import re, csv, collections
import tia_core as c

LEVELS = ['errore', 'termine', 'mancante', 'verifica']
YES = {'s', 'si', 'sì', 'x', 'y', 'yes', '1', 'ok'}
SEP_NL = '\n'

def _nl(s):
    return (s or '').replace('\r\n', '\n').replace('\r', '\n')

def _norm(s):
    """Per confronti: senza segnaposto, spazi e maiuscole."""
    return re.sub(r'\s+', ' ', c.FIELD_RE.sub('{}', _nl(s))).strip().lower()

def _langmatch(spec, lang):
    spec = spec.strip()
    return spec == '*' or spec == lang or spec == lang.split('-')[0]

# ---------------------------------------------------------------- errori noti
class KnownErrors:
    def __init__(self, path=None):
        self.typos = []     # (lingua, regex, giusto)
        self.terms = []     # (lingua, regex it, regex errata, suggerita, nota)
        if path:
            self.load(path)

    def load(self, path):
        import os
        if not os.path.exists(path):
            return
        for ln in open(path, encoding='utf-8'):
            ln = ln.strip()
            if not ln or ln.startswith('#'):
                continue
            p = [x.strip() for x in ln.split(';')]
            try:
                if p[0] == 'refuso' and len(p) >= 4:
                    rx = re.compile(r'(?<!\w)' + r'\s+'.join(re.escape(w) for w in p[2].split()) + r'(?!\w)', re.I)
                    self.typos.append((p[1], rx, p[3]))
                elif p[0] == 'termine' and len(p) >= 5:
                    self.terms.append((p[1], re.compile(p[2], re.I), re.compile(p[3], re.I), p[4], p[5] if len(p) > 5 else ''))
            except re.error:
                continue

def _case_like(sample, new):
    if len(sample) > 1 and sample.isupper():
        return new.upper()
    if sample[:1].isupper() and not sample.isupper():
        return new[:1].upper() + new[1:]
    return new

# ---------------------------------------------------------------- contraddizioni (antonimi)
# (regex italiano, {lingua: regex}) - se l'originale ha un concetto e non il suo opposto, la traduzione non deve avere
# l'opposto senza il concetto.
def _pair(a_it, b_it, a, b):
    return ((a_it, a), (b_it, b))
ANTONYMS = [
    _pair(r'massim|\bmax', r'minim|\bmin\b',
          {'en': r'\bmax', 'de': r'\bmax', 'es': r'm[aá]x', 'pl': r'maks|\bmax'},
          {'en': r'\bmin', 'de': r'\bmin', 'es': r'm[ií]n', 'pl': r'\bmin'}),
    _pair(r'\bapri|apert|apertura', r'\bchiud|chius|chiusura',
          {'en': r'\bopen', 'de': r'[öo]ffn', 'es': r'\babr', 'pl': r'otw'},
          {'en': r'\bclos|\bshut', 'de': r'schlie', 'es': r'\bcerr', 'pl': r'zamk'}),
    _pair(r'anterior|\bavanti\b|\bfronte', r'posterior|indietro|\bretro',
          {'en': r'front|\bforw', 'es': r'delanter|adelante|anterior', 'pl': r'przedn|\bprz[oó]d', 'de': r'vorn|vorw'},
          {'en': r'\brear|\bback', 'es': r'trasero|posterior|atr[aá]s', 'pl': r'tyln|\bty[lł]', 'de': r'hinten|r[üu]ckw'}),
    _pair(r'(?<!dis)abilit', r'disabilit|inibi',
          {'en': r'(?<!dis)\benabl', 'es': r'(?<!des)habilit', 'pl': r'(?<!wy)[łl][aą]cz'},
          {'en': r'disabl|inhibit', 'es': r'deshabilit|inhibi', 'pl': r'wy[łl][aą]cz|blokad'}),
    _pair(r'destr', r'sinistr',
          {'en': r'\bright', 'es': r'derech', 'pl': r'praw', 'de': r'rechts'},
          {'en': r'\bleft', 'es': r'izquierd', 'pl': r'\blew', 'de': r'links'}),
    _pair(r'(?<!s)caric', r'scaric',
          {'en': r'(?<!un)load', 'es': r'(?<!des)carg', 'pl': r'(?<!roz)(?<!wy)[łl]ad', 'de': r'(?<!ent)lad'},
          {'en': r'unload|discharg', 'es': r'descarg', 'pl': r'roz[łl]ad|wy[łl]ad', 'de': r'entlad'}),
    _pair(r'\bsal[ei]|\bsalita|sollev|\balza', r'disces|abbass|\bgi[uù]\b',
          {'en': r'\bup\b|\brais|\blift', 'es': r'\bsub|\belev', 'pl': r'\bw g[oó]r|podnie', 'de': r'\bauf\b|heben'},
          {'en': r'\bdown\b|\blower', 'es': r'\bbaj', 'pl': r'\bw d[oó][łl]|opu[sś]', 'de': r'\bab\b|senk'}),
]

def _antonym_issue(src, tr, lang):
    lg = lang.split('-')[0]
    s = c.FIELD_RE.sub('', src).lower()
    t = c.FIELD_RE.sub('', tr).lower()
    for pa, pb in ANTONYMS:
        for (x_it, x), (y_it, y) in ((pa, pb), (pb, pa)):
            if lg not in x or lg not in y:
                continue
            if re.search(x_it, s) and not re.search(y_it, s):
                if re.search(y[lg], t) and not re.search(x[lg], t):
                    return 'contraddice l\'originale (l\'italiano ha "%s", la traduzione l\'opposto)' % re.search(x_it, s).group(0)
    return None

# ---------------------------------------------------------------- controlli sulla singola cella
def _issues(src, cur, lang, ctx):
    """Lista di (livello, problema, proposta|None, fonte)."""
    out = []
    s, t = _nl(src).strip(), _nl(cur).strip()
    # strutturali, uguali a quelli applicati all'output del modello (senza il limite HMI, trattato a parte)
    for m in c.check_translation(s, t, lang, ctx['glossary']):
        if m.startswith('glossario:'):
            out.append(('termine', m, None, ''))
        elif m.startswith('molto piu') or m.startswith('numero di righe'):   # nei testi HMI le righe si adattano all'oggetto
            out.append(('verifica', m, None, ''))
        else:
            out.append(('errore', m, None, ''))
    # testo predefinito TIA rimasto
    if t in ('ON', 'OFF', 'On', 'Off', 'Text') and s not in ('ON', 'OFF', 'On', 'Off', 'Text'):
        out.append(('errore', 'rimasto il testo predefinito di TIA ("%s")' % t, None, ''))
    # numeri e lato A/B
    ns = collections.Counter(re.findall(r'\d+', c.FIELD_RE.sub('', s)))
    nt = collections.Counter(re.findall(r'\d+', c.FIELD_RE.sub('', t)))
    if not re.search(r'(?<!\w)zero(?!\w)', s, re.I):      # "a 0" tradotto "to Zero" (ma "Premere ZERO" non e' un numero)
        for w in re.findall(r'(?<!\w)(?:zero|cero|null)(?!\w)', c.FIELD_RE.sub('', t), re.I):
            nt['0'] += 1
    if ns != nt and ns:
        out.append(('errore', 'numeri diversi dall\'originale (originale: %s; traduzione: %s)' % (
            ' '.join(sorted(ns.elements())) or '-', ' '.join(sorted(nt.elements())) or '-'), None, ''))
    ms = None if re.search(r'\b[A-Z]\s*/\s*[A-Z]\b', s) else re.search(r'\b(?:lato|side|lado|seite|strona)\s+([A-Z])\b', s, re.I)   # "Lato A/B": nessun lato
    mt = re.search(r'\b(?:lato|side|lado|seite|strona)\s+([A-Z])\b|\b([A-Z])\s+(?:side|lado|seite)\b', t, re.I)
    if ms and mt and ms.group(1).upper() != (mt.group(1) or mt.group(2)).upper():
        out.append(('errore', 'lato diverso dall\'originale (%s)' % ms.group(1).upper(), None, ''))
    # contraddizioni
    a = _antonym_issue(s, t, lang)
    if a:
        out.append(('errore', a, None, ''))
    # errori noti: refusi (proposta sicura) e termini
    fixed = t
    typo_hit = False
    for l, rx, giusto in ctx['known'].typos:
        if _langmatch(l, lang) and rx.search(fixed):
            fixed = rx.sub(lambda m: _case_like(m.group(0), giusto), fixed)
            typo_hit = True
    if typo_hit:
        out.append(('errore', 'refuso noto', _restore_ws(cur, fixed), 'refuso'))
    for l, rx_it, rx_bad, sug, note in ctx['known'].terms:
        if _langmatch(l, lang) and rx_it.search(s) and rx_bad.search(t):
            prop = rx_bad.sub(lambda m: _case_like(m.group(0), sug), t, count=1)
            out.append(('termine', 'termine sbagliato: %s' % (note or sug), _restore_ws(cur, prop), 'termine'))
    return out

def _restore_ws(orig, new):
    lead, _, trail = c.split_ws(_nl(orig))
    return lead + new + trail

# ---------------------------------------------------------------- audit del progetto
def audit_project(project, targets, glossary, tm, known, hmi=None, only_cats=None, exclude_graph=False):
    """Ritorna la lista di gruppi: dict(livello, lingua, cat, righe[], src, cur, problemi[], proposta, fonte, applica)."""
    ctx = {'glossary': glossary, 'known': known}
    en_lang = next((l for l in targets if l.startswith('en')), None)
    rows = []
    for r in project.rows:
        core = r['src'].strip()
        if c.class_row_info(r['vp'], r['cat']) or c.is_copy_rule(core):
            continue
        if exclude_graph and c.is_graph_row(r['vp']):
            continue
        if only_cats and not any(x.lower() in r['cat'].lower() for x in only_cats):
            continue
        rows.append(r)
    sources = {_norm(r['src']): r['src'] for r in project.rows}
    cells = []     # (riga, lingua, livello, problema, proposta, fonte)
    variants = collections.defaultdict(lambda: collections.defaultdict(list))   # (core,lang,hmiflag) -> testo -> righe
    shared = collections.defaultdict(lambda: collections.defaultdict(list))     # (lang, norm cur) -> norm src -> righe
    for r in rows:
        core = r['src'].strip()
        row_hmi = c.is_hmi_row(r['cat'], r['vp'], hmi)
        for l in targets:
            cur = _nl(r['cur'].get(l, ''))
            if c.is_placeholder(cur, _nl(r['src'])):
                prop = tm.get(core, l + c.HMI_SUFFIX) if row_hmi else None
                if prop is None and not (row_hmi and (not tm.get(core, l) or c.hmi_overflow(core, tm.get(core, l), hmi, l))):
                    prop = tm.get(core, l)
                if prop == '=':
                    prop = core
                cells.append((r, l, 'mancante', 'traduzione mancante', _restore_ws(r['src'], prop) if prop else None,
                              'memoria' if prop else ''))
                continue
            iss = _issues(r['src'], cur, l, ctx)
            if row_hmi:
                ov = c.hmi_overflow(core, cur.strip(), hmi, l)
                if ov:
                    iss.append(('verifica', 'troppo lunga per HMI (%d car., max %d)' % ov, None, ''))
            # uguale all'inglese (lingua non inglese): probabilmente non tradotta
            if en_lang and l != en_lang and not l.startswith('en'):
                en = _nl(r['cur'].get(en_lang, ''))
                if en and not c.is_placeholder(en, _nl(r['src'])) and _norm(en) == _norm(cur) \
                        and len(c.FIELD_RE.sub('', cur).strip()) > 3 and c.has_letters(cur):
                    iss.append(('verifica', 'uguale alla traduzione %s: forse non tradotta' % en_lang, None, ''))
            # la traduzione e' un altro testo italiano del progetto
            nc = _norm(cur)
            if nc != _norm(r['src']) and nc in sources and len(nc) >= 8 and nc not in _norm(r['src']):
                iss.append(('verifica', 'la traduzione e\' un altro testo italiano del progetto', None, ''))
            # proposta dalla memoria (solo se "rivisto" da te)
            suf = c.HMI_SUFFIX if row_hmi else ''
            e = tm.entries.get(core, {}).get(l + suf)
            if iss and e and e['s'] == 'rivisto' and _norm(e['t']) != _norm(cur):
                iss.append(('verifica', 'la memoria ha una traduzione rivista diversa', _restore_ws(r['src'], e['t']), 'memoria'))
            for (lv, pr, pp, fo) in iss:
                cells.append((r, l, lv, pr, pp, fo))
            variants[(core, l, row_hmi)][_norm(cur)].append((r['row'], cur))
            shared[(l, nc)][_norm(r['src'])].append(r['row'])
    # coerenza interna: stesso italiano tradotto in modi diversi
    for (core, l, hm), d in variants.items():
        if len(d) > 1 and not hm:
            best = max(d, key=lambda k: len(d[k]))
            for k, lst in d.items():
                if k == best:
                    continue
                for rown, cur in lst:
                    r = next(x for x in rows if x['row'] == rown)
                    cells.append((r, l, 'verifica', 'incoerente: le altre %d righe con lo stesso italiano hanno "%s"' % (
                        len(d[best]), d[best][0][1].replace('\n', ' / ')[:60]), None, ''))
    # raggruppa le celle identiche
    byrow = collections.defaultdict(list)
    for (r, l, lv, pr, pp, fo) in cells:
        byrow[(r['row'], l)].append((lv, pr, pp, fo, r))
    groups = collections.OrderedDict()
    for (rown, l), items in byrow.items():
        r = items[0][4]
        lv = min((i[0] for i in items), key=LEVELS.index)
        prob = ' | '.join(dict.fromkeys(i[1] for i in items))
        props = [(i[2], i[3]) for i in items if i[2]]
        # una proposta sola, la piu' affidabile: refuso > memoria > termine
        order = {'refuso': 0, 'memoria': 1, 'termine': 2, '': 3}
        props.sort(key=lambda x: order.get(x[1], 3))
        prop, fonte = props[0] if props else ('', '')
        cur = _nl(r['cur'].get(l, ''))
        key = (lv, l, r['cat'], r['src'], cur, prob, prop, fonte)
        groups.setdefault(key, []).append(rown)
    out = []
    for (lv, l, cat, src, cur, prob, prop, fonte), rr in groups.items():
        applica = 's' if fonte in ('refuso',) else ''
        out.append({'livello': lv, 'lingua': l, 'cat': cat.strip('<>'), 'righe': sorted(rr), 'src': src.strip(),
                    'cur': cur, 'problemi': prob, 'proposta': prop, 'fonte': fonte, 'applica': applica})
    out.sort(key=lambda g: (LEVELS.index(g['livello']), g['lingua'], g['cat'], g['src'], g['righe'][0]))
    return out

def export_audit(path, groups):
    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['livello', 'lingua', 'categoria', 'righe', 'n', 'italiano', 'traduzione attuale', 'problema',
                    'proposta', 'fonte proposta', 'applica'])
        for g in groups:
            w.writerow([g['livello'], g['lingua'], g['cat'], ', '.join(map(str, g['righe'])), len(g['righe']),
                        g['src'], g['cur'], g['problemi'], g['proposta'], g['fonte'], g['applica']])

def summary(groups):
    cnt = collections.Counter()
    for g in groups:
        cnt[g['livello']] += len(g['righe'])
    return cnt

# ---------------------------------------------------------------- applicazione delle correzioni
def read_audit_corrections(path):
    """Righe del CSV da applicare: [(lingua, [righe], traduzione attuale, proposta)]."""
    out = []
    with open(path, newline='', encoding='utf-8-sig') as f:
        rd = csv.DictReader(f, delimiter=';')
        for row in rd:
            if (row.get('applica') or '').strip().lower() not in YES or not (row.get('proposta') or '').strip('\r\n '):
                continue
            rr = [int(x) for x in re.findall(r'\d+', row.get('righe') or '')]
            out.append((row['lingua'].strip(), rr, _nl(row.get('traduzione attuale', '')), _nl(row['proposta'])))
    return out

def plan_corrections(project, corrections):
    """Ritorna (edits {(riga, colonna): testo}, learn [(core, lingua, testo, riga)], scartate [(riga, lingua, motivo)]).
    Una cella si cambia solo se contiene ancora il testo controllato e la proposta ha gli stessi segnaposto e righe."""
    byrow = {r['row']: r for r in project.rows}
    edits, learn, bad = {}, [], []
    for lang, rows, attuale, prop in corrections:
        if lang not in project.cols:
            bad.append((rows[:1] and rows[0] or 0, lang, 'lingua non presente nell\'Excel')); continue
        for n in rows:
            r = byrow.get(n)
            if not r:
                bad.append((n, lang, 'riga non trovata')); continue
            cur = _nl(r['cur'].get(lang, ''))
            if not c.is_placeholder(cur, _nl(r['src'])) and cur.strip() != attuale.strip():
                bad.append((n, lang, 'la cella e\' cambiata dopo il controllo')); continue
            p = prop.strip('\r\n')
            if c.FIELD_RE.findall(p) != c.FIELD_RE.findall(_nl(r['src'])):
                bad.append((n, lang, 'segnaposto <field> diversi dall\'originale')); continue
            if p.count('\n') != _nl(r['src']).strip().count('\n'):
                bad.append((n, lang, 'numero di righe diverso dall\'originale')); continue
            base = cur if not c.is_placeholder(cur, _nl(r['src'])) else r['src']
            lead, _, trail = c.split_ws(_nl(base))
            edits[(n, project.cols[lang])] = lead + p.strip() + trail
            learn.append((r['src'].strip(), lang, p.strip(), r))
    return edits, learn, bad
