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

- GRAPH: i nomi di passi e transizioni (riconosciuti dal percorso: liste di sistema SYSTEM_StepNames / SYSTEM_TransitionNames) non vengono tradotti (casella nella finestra, `escludi_graph` in config).
- Per progetti grandi: spunta "Col modello traduci solo i testi di allarme" (categorie in `config.json` → `categorie_allarme`; CLI: `--solo-allarmi`).

## Lunghezza dei testi HMI
Nell'Excel non c'è la larghezza degli oggetti: l'unico riferimento è la lunghezza dell'originale italiano.
Per le righe HMI (categorie in `categorie_hmi`, default `HMI screen` e `HMI runtime`; i commenti HMI non si vedono a video e non hanno limite) ogni riga della traduzione può superare l'italiano di al massimo `hmi_soglia_percentuale` % (default 20), con un margine minimo di `hmi_margine_minimo` caratteri (default 2). I segnaposto `<field/>` non contano.
- Il modello riceve i limiti per riga e l'istruzione di abbreviare; se sfora, il testo va nel residuo con motivo `HMI | … troppo lunga per HMI (N car., max M)`.
- Se la memoria ha solo una traduzione normale che sfora, per quella riga HMI non viene usata e il testo viene riscritto in forma breve.
- Le versioni brevi sono salvate nella memoria sotto `<lingua>@hmi`, **accanto** a quelle normali: un allarme con lo stesso testo italiano usa sempre la traduzione normale.
- Nel CSV del residuo le righe HMI hanno il motivo che inizia con `HMI`: **non cambiare la colonna `motivo`**, così "Importa correzioni" le salva come versioni brevi.
- Impostazioni… permette di cambiare soglia e margine e di spegnere il controllo (`controllo_lunghezza_hmi`).
- Soglie per lingua in `config.json` → `hmi_soglia_per_lingua` (default `{"pl-PL": 30}`: il polacco è mediamente più lungo dell'italiano; le altre lingue usano `hmi_soglia_percentuale`). Taratura fatta su G2290 (HMI già tradotto a mano e adattato): rapporto di lunghezza mediano/90° percentile en-US 0,90/1,00, es-MX 0,93/1,14, pl-PL 1,11/1,25.
- "Impara da Excel già tradotto" salva le righe HMI come versioni brevi (`@hmi`), non nella memoria normale: i testi adattati all'oggetto non finiscono negli allarmi.

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
