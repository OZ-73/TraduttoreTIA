# -*- coding: utf-8 -*-
"""Riga di comando opzionale. Esempi:
  python tia_cli.py progetto.xlsx --out progetto_tradotto.xlsx --traduci --residuo residuo.csv
  python tia_cli.py progetto.xlsx --importa residuo_corretto.csv --out progetto_tradotto.xlsx
  python tia_cli.py progetto.xlsx --impara altro_progetto_tradotto.xlsx"""
import argparse
import tia_core as c

def main():
    ap = argparse.ArgumentParser(description='Traduttore TIA')
    ap.add_argument('excel')
    ap.add_argument('--out'); ap.add_argument('--traduci', action='store_true')
    ap.add_argument('--residuo'); ap.add_argument('--importa'); ap.add_argument('--impara')
    ap.add_argument('--solo-allarmi', action='store_true')
    ap.add_argument('--lingue', help='es. en-US,de-DE (default: tutte)')
    a = ap.parse_args()
    s = c.Session()
    s.open(a.excel)
    s.only_alarms = a.solo_allarmi
    if a.lingue:
        s.targets = a.lingue.split(',')
    if a.impara:
        o = c.Project(a.impara, s.cfg); print('Appresi', c.learn_from_project(o, s.tm, o.targets)); s.tm.save()
    if a.importa:
        print('Importate', c.import_corrections(a.importa, s.tm, s.targets)); s.tm.save()
    e, p, st, notes = s.plan(); print(st, 'da tradurre:', len(p))
    if a.traduci:
        print(s.translate_pending(lambda d, n, el, t: print('%d/%d %ds' % (d, n, el))))
    if a.residuo:
        c.export_residual(a.residuo, s.review, s.targets); print('Residuo:', len(s.review))
    if a.out:
        print(s.save_excel(a.out))

if __name__ == '__main__':
    main()
