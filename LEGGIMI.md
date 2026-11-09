# Traduttore TIA (motore offline)

Requisiti: Python 3 con tkinter (incluso nell'installer standard di Windows) e `pip install openpyxl`.
LM Studio con il server attivo (Developer → Status: Running) e il modello caricato (default `openai/gpt-oss-120b`, vedi Impostazioni).

## Flusso
1. In TIA: esporta i testi in Excel. Nell'Excel devono esserci le colonne delle lingue di destinazione (es. en-US, de-DE) già aggiunte al progetto.
2. Avvia con `avvia.bat` (o `python tia_gui.py`).
3. **1. Apri Excel** → scegli le lingue → **2. Analizza**: applica memoria, regole (codici/identificatori), classi di allarme con sigle.
4. **3. Traduci i nuovi testi**: il modello traduce solo ciò che la memoria non copre. Le traduzioni che superano i controlli entrano nella memoria; le altre (e un campione del 10%) vanno in revisione.
5. **4. Esporta residuo (CSV)** → mandalo a Claude in chat → ricevi il CSV corretto → **5. Importa correzioni** (entrano in memoria come "rivisto", hanno la priorità).
6. **6. Salva Excel tradotto** (con verifica cella per cella) → importa in TIA → archivia.

- GRAPH: i nomi di passi e transizioni (testi che iniziano con S, T, TR, SR) non vengono tradotti (casella nella finestra, `escludi_graph` in config).
- Per progetti grandi: spunta "Col modello traduci solo i testi di allarme" (categorie in `config.json` → `categorie_allarme`; CLI: `--solo-allarmi`).

## File
- `memoria_traduzione.json`: testi già tradotti (parte dal progetto Comipont); cresce a ogni uso. Tienila da parte: è il valore del sistema.
- `glossario.txt`: termini fissi (modificabile a mano: `italiano | inglese | tedesco`, `=` invariato).
- `classi_allarme.json`: nomi e sigle delle classi per lingua.
- `config.json`: impostazioni (anche da pulsante).

## Regole
- Scrive solo nelle celle vuote, uguali all'originale, "Text" o con soli segnaposto: non sovrascrive mai traduzioni esistenti.
- Le etichette già presenti ma errate (es. schermate HMI con testo italiano diverso dall'originale) non vengono toccate: correggile a mano o dall'Excel.
- Testi di sistema Siemens (~260) restano com'erano.
- Riga di comando opzionale: `python tia_cli.py progetto.xlsx --traduci --residuo residuo.csv --out out.xlsx` (vedi `tia_cli.py`).
