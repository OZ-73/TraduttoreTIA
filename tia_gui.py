# -*- coding: utf-8 -*-
"""Traduttore TIA - finestra con i pulsanti (tkinter)."""
import os, sys, threading, queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tia_core as core

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Traduttore TIA - motore offline (LM Studio)')
        self.geometry('900x640')
        try:
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            self.iconbitmap(os.path.join(base, 'icona.ico'))
        except Exception:
            pass
        self.q = queue.Queue()
        self.s = core.Session(log=lambda m: self.q.put(('log', m)))
        self.lang_vars = {}
        self.busy = False
        self._build()
        self.after(100, self._pump)
        self.log('Pronto. TM: %d testi, glossario: %d termini, classi: %d.' %
                 (len(self.s.tm.entries), len(self.s.glossary.terms), len(self.s.classes)))

    # ---------------- interfaccia
    def _build(self):
        top = ttk.Frame(self, padding=8); top.pack(fill='x')
        ttk.Button(top, text='1. Apri Excel TIA…', command=self.on_open).pack(side='left')
        self.lbl_file = ttk.Label(top, text='  (nessun file)'); self.lbl_file.pack(side='left')
        ttk.Button(top, text='Impostazioni…', command=self.on_settings).pack(side='right')

        self.lang_frame = ttk.LabelFrame(self, text='Lingue da compilare', padding=6)
        self.lang_frame.pack(fill='x', padx=8)

        self.var_alarms = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text='Col modello traduci solo HMI e allarmi (commenti di blocco e di tag restano com\'erano; la memoria vale sempre)',
                        variable=self.var_alarms, command=self.refresh_stats).pack(anchor='w', padx=10)

        self.var_graph = tk.BooleanVar(value=bool(self.s.cfg.get('escludi_graph', True)))
        ttk.Checkbutton(self, text='Non tradurre i nomi di passi e transizioni GRAPH (liste di sistema StepNames / TransitionNames)',
                        variable=self.var_graph, command=self.on_graph).pack(anchor='w', padx=10)

        self.lbl_stats = ttk.Label(self, text='', justify='left'); self.lbl_stats.pack(fill='x', padx=10, pady=4)

        bf = ttk.LabelFrame(self, text='Operazioni', padding=6); bf.pack(fill='x', padx=8)
        self.buttons = []
        spec = [('2. Analizza (applica memoria e regole)', self.on_analyze),
                ('3. Traduci i nuovi testi con il modello', self.on_translate),
                ('Interrompi', self.on_stop),
                ('4. Esporta residuo da rivedere (CSV)', self.on_export),
                ('5. Importa correzioni (CSV)', self.on_import),
                ('6. Salva Excel tradotto…', self.on_save),
                ('Impara da Excel già tradotto…', self.on_learn),
                ('Controllo coerenza', self.on_consistency)]
        for i, (t, f) in enumerate(spec):
            b = ttk.Button(bf, text=t, command=f)
            b.grid(row=i // 2, column=i % 2, sticky='ew', padx=4, pady=3)
            self.buttons.append(b)
            if t == 'Interrompi':
                self.btn_stop = b
        bf.columnconfigure(0, weight=1); bf.columnconfigure(1, weight=1)

        self.pb = ttk.Progressbar(self, mode='determinate'); self.pb.pack(fill='x', padx=8, pady=6)
        self.lbl_prog = ttk.Label(self, text=''); self.lbl_prog.pack(fill='x', padx=10)
        self.txt = tk.Text(self, height=14, wrap='word'); self.txt.pack(fill='both', expand=True, padx=8, pady=6)

    def log(self, m):
        self.txt.insert('end', m + '\n'); self.txt.see('end')

    def _pump(self):
        try:
            while True:
                k, v = self.q.get_nowait()
                if k == 'log':
                    self.log(v)
                elif k == 'prog':
                    d, n, el, tok = v
                    self.pb['maximum'] = max(n, 1); self.pb['value'] = d
                    eta = el / d * (n - d) if d else 0
                    self.lbl_prog.config(text='%d / %d testi  -  trascorsi %d s  -  restano ~%d s  -  %d token' % (d, n, el, eta, tok))
                elif k == 'done':
                    self.busy = False
                    if v:
                        self.log(v)
                    self.refresh_stats()
        except queue.Empty:
            pass
        self.after(150, self._pump)

    def on_graph(self):
        self.s.cfg['escludi_graph'] = self.var_graph.get()
        core.save_config(self.s.cfg)
        self.refresh_stats()

    def targets(self):
        return [l for l, v in self.lang_vars.items() if v.get()]

    def need_file(self):
        if not self.s.project:
            messagebox.showinfo('Traduttore TIA', 'Apri prima un Excel TIA.')
            return False
        if self.busy:
            messagebox.showinfo('Traduttore TIA', 'Operazione in corso.')
            return False
        self.s.targets = self.targets()
        self.s.only_alarms = self.var_alarms.get()
        return True

    def run_bg(self, fn, msg=None):
        self.busy = True
        def w():
            try:
                r = fn()
                self.q.put(('done', msg(r) if callable(msg) else msg))
            except Exception as e:
                self.q.put(('done', 'ERRORE: %s' % e))
        threading.Thread(target=w, daemon=True).start()

    def refresh_stats(self):
        if not self.s.project:
            return
        try:
            self.s.targets = self.targets()
            self.s.only_alarms = self.var_alarms.get()
            e, p, st, notes = self.s.plan()
            self.lbl_stats.config(text='Righe: %d   Celle da compilare: %d\n'
                                       'Coperte da memoria: %d   regole: %d   classi allarme: %d   ancora da tradurre (testi distinti): %d'
                                       '   (di cui HMI da scrivere brevi: %d)\n'
                                       'In revisione (bozze/campione): %d' %
                                  (st['righe'], st['celle_da_compilare'], st['tm'], st['regole'], st['classi'],
                                   len(p) + len(self.s.pending_hmi), len(self.s.pending_hmi), len(self.s.review)))
        except Exception as ex:
            self.lbl_stats.config(text='Errore analisi: %s' % ex)

    # ---------------- azioni
    def on_open(self):
        p = filedialog.askopenfilename(title='Excel esportato da TIA', filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        try:
            pr = self.s.open(p)
        except Exception as e:
            messagebox.showerror('Traduttore TIA', str(e)); return
        self.lbl_file.config(text='  ' + os.path.basename(p))
        for w in self.lang_frame.winfo_children():
            w.destroy()
        self.lang_vars = {}
        for l in pr.targets:
            v = tk.BooleanVar(value=True); self.lang_vars[l] = v
            ttk.Checkbutton(self.lang_frame, text=l, variable=v, command=self.refresh_stats).pack(side='left', padx=8)
        self.log('Aperto %s: %d righe, lingue: %s' % (os.path.basename(p), len(pr.rows), ', '.join(pr.targets)))
        self.refresh_stats()

    def on_analyze(self):
        if self.need_file():
            self.s.reload_resources(); self.refresh_stats()
            e, p, st, notes = self.s.plan()
            self.log('Analisi: %s' % st)
            for n in notes[:20]:
                self.log('  nota: ' + n)
            self.log('Testi da tradurre col modello: %d (HMI da scrivere brevi: %d)' % (len(p) + len(self.s.pending_hmi), len(self.s.pending_hmi)))

    def on_translate(self):
        if not self.need_file():
            return
        self.s.reload_resources()
        n = len(self.s.plan()[1]) + len(self.s.pending_hmi)
        if n == 0:
            messagebox.showinfo('Traduttore TIA', 'Niente da tradurre: tutto è coperto dalla memoria.'); return
        if not messagebox.askyesno('Traduttore TIA', 'Tradurre %d testi con il modello "%s"?\n(LM Studio deve avere il server attivo.)' % (n, self.s.cfg['modello'])):
            return
        self.log('Traduzione di %d testi…' % n)
        self.run_bg(lambda: self.s.translate_pending(lambda d, n, el, tok: self.q.put(('prog', (d, n, el, tok)))),
                    lambda r: 'Finito: %d testi tradotti, %d da rivedere (bozze + campione).' % r)

    def on_stop(self):
        self.s.stop = True

    def on_export(self):
        if not self.need_file():
            return
        if not self.s.review:
            messagebox.showinfo('Traduttore TIA', 'Nessun testo da rivedere.'); return
        p = filedialog.asksaveasfilename(defaultextension='.csv', initialfile='residuo.csv', filetypes=[('CSV', '*.csv')])
        if p:
            core.export_residual(p, self.s.review, self.targets())
            self.log('Esportati %d testi in %s. Mandali a Claude in chat per la revisione.' % (len(self.s.review), p))

    def on_import(self):
        if not self.need_file():
            return
        p = filedialog.askopenfilename(title='CSV corretto', filetypes=[('CSV', '*.csv')])
        if not p:
            return
        n = core.import_corrections(p, self.s.tm, self.targets())
        self.s.tm.save()
        self.s.prune_review()
        self.log('Importate %d traduzioni nella memoria (stato "rivisto").' % n)
        self.refresh_stats()

    def on_save(self):
        if not self.need_file():
            return
        base = os.path.splitext(os.path.basename(self.s.project.path))[0]
        p = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile=base + '_tradotto.xlsx', filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        if os.path.abspath(p) == os.path.abspath(self.s.project.path):
            messagebox.showerror('Traduttore TIA', 'Scegli un nome diverso dall\'originale.'); return
        self.run_bg(lambda: self.s.save_excel(p),
                    lambda r: 'Salvato. Celle scritte: %(scritte)d, modificate: %(modificate)d, anomalie: %(anomalie)d, ancora da tradurre: %(ancora_da_tradurre)d' % r)

    def on_learn(self):
        if not self.need_file():
            return
        p = filedialog.askopenfilename(title='Excel già tradotto', filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        try:
            other = core.Project(p, self.s.cfg)
            n = core.learn_from_project(other, self.s.tm, [l for l in other.targets], core.hmi_settings(self.s.cfg))
            self.s.tm.save()
            self.log('Appresi %d nuovi testi da %s.' % (n, os.path.basename(p)))
            self.refresh_stats()
        except Exception as e:
            messagebox.showerror('Traduttore TIA', str(e))

    def on_consistency(self):
        if not self.need_file():
            return
        rep = core.consistency_report(self.s.project, self.targets())
        self.log('Coerenza: %d testi con traduzioni diverse.' % len(rep))
        for c, l, v in rep[:40]:
            self.log('  [%s] %s -> %s' % (l, c[:50], ' | '.join(v)))

    def on_settings(self):
        w = tk.Toplevel(self); w.title('Impostazioni'); w.grab_set()
        keys = [('url', 'Indirizzo server'), ('modello', 'Modello'), ('dimensione_lotto', 'Testi per richiesta'),
                ('percentuale_campione', 'Campione di controllo %'), ('reasoning_effort', 'Reasoning (low/medium/high)'),
                ('esempi_nel_prompt', 'Esempi dalla memoria nel prompt'), ('timeout_s', 'Timeout (s)'),
                ('hmi_soglia_percentuale', 'HMI: max. % di caratteri in più dell\'originale'),
                ('hmi_margine_minimo', 'HMI: margine minimo (caratteri)')]
        vars_ = {}
        for i, (k, lab) in enumerate(keys):
            ttk.Label(w, text=lab).grid(row=i, column=0, sticky='w', padx=6, pady=3)
            v = tk.StringVar(value=str(self.s.cfg.get(k, ''))); vars_[k] = v
            ttk.Entry(w, textvariable=v, width=45).grid(row=i, column=1, padx=6)
        bv = tk.BooleanVar(value=bool(self.s.cfg.get('applica_bozze_da_rivedere')))
        ttk.Checkbutton(w, text='Scrivi nell\'Excel anche le bozze da rivedere', variable=bv).grid(row=len(keys), column=0, columnspan=2, sticky='w', padx=6)
        hv = tk.BooleanVar(value=bool(self.s.cfg.get('controllo_lunghezza_hmi', True)))
        ttk.Checkbutton(w, text='Controlla la lunghezza dei testi HMI (categorie in config.json → categorie_hmi)',
                        variable=hv).grid(row=len(keys) + 2, column=0, columnspan=2, sticky='w', padx=6)
        def ok():
            for k, v in vars_.items():
                old = self.s.cfg.get(k)
                val = v.get().strip()
                try:
                    if isinstance(old, bool):
                        continue
                    self.s.cfg[k] = int(float(val)) if isinstance(old, int) else float(val) if isinstance(old, float) else val
                except ValueError:
                    messagebox.showerror('Impostazioni', 'Valore non valido per %s' % k); return
            self.s.cfg['applica_bozze_da_rivedere'] = bv.get()
            self.s.cfg['controllo_lunghezza_hmi'] = hv.get()
            core.save_config(self.s.cfg); w.destroy(); self.refresh_stats()
        ttk.Button(w, text='Salva', command=ok).grid(row=len(keys) + 3, column=1, sticky='e', padx=6, pady=8)

def main():
    App().mainloop()

if __name__ == '__main__':
    main()
