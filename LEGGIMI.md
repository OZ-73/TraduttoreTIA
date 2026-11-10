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
- Per progetti grandi: spunta "Col modello traduci solo HMI e allarmi" (CLI: `--solo-allarmi`). Col modello si traducono solo le righe la cui **categoria** (colonna Category) contiene una voce di `config.json` → `categorie_da_tradurre` (default `Alarm`, `TextListText`, `ProDiag`, `HMI`): ProgramAlarm (`<FormattedUserAlarmTextCategory>`), allarmi HMI (`<Alarm>`), classi di allarme, elenchi di testo PLC (`<TextListTextCategoryData>`) e tutte le categorie `<HMI …>`. Commenti di blocco (`<BlockCommentCategoryData>`) e di tag (`<TagCommentCategoryData>`) restano com'erano; la memoria si applica comunque a tutte le righe. Il percorso (ViewPath) non conta più.

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

## Controllo delle traduzioni già presenti
Per un Excel già tradotto (a mano, con Google Translate, da un fornitore): **Controlla traduzioni esistenti (CSV)…** (CLI: `--controlla file.csv`). Rispetta lingue scelte, ambito "solo HMI e allarmi" e GRAPH. Non usa il modello e non cambia nulla: produce un CSV, una riga per gruppo di celle identiche, con le righe Excel complete.
- **livello `errore`**: segnaposto `<field>` diversi, numeri diversi dall'originale (es. FL2 → FL1, SP32803 → SP12803), lato A/B diverso, testo predefinito TIA rimasto (ON/OFF/Text), parole italiane rimaste, contraddizioni (max/min, apri/chiudi, avanti/indietro, abilita/disabilita, destra/sinistra, carico/scarico, su/giù), refusi noti.
- **`termine`**: glossario non rispettato e falsi amici/termini tecnici sbagliati (`errori_noti.txt`, es. bussola = bushing, non compass).
- **`mancante`**: celle vuote o ancora col testo originale; la proposta viene dalla memoria se c'è.
- **`verifica`**: uguale alla traduzione inglese (forse non tradotta), stesso italiano tradotto in modi diversi, traduzione che coincide con un altro testo italiano, numero di righe diverso, troppo lunga per HMI (con testi già adattati a mano è normale che ce ne siano pochi).
- Il controllo trova gli errori **di forma e di vocabolario noto**, non quelli di senso (una frase fluente ma sbagliata passa): per quelli manda il CSV a Claude in chat. Nei test su G2290 ha segnalato 102 su 111 righe con errori certi trovati a mano; le 9 mancate erano semantiche.
- `errori_noti.txt` è modificabile a mano (formato nel file): `refuso` = sostituzione sicura (applica = s preimpostato), `termine` = resa sbagliata con suggerimento da confermare. Aggiungi i tuoi errori ricorrenti.
- Nel controllo "uguale alla traduzione inglese" e il glossario vengono confrontati con ciò che c'è nel progetto: un termine di progetto diverso dal glossario (es. Avanzatore: Advancer nel progetto, Feeder nel glossario) esce come `termine`; decidi se cambiare il glossario o il progetto.

**Applicare le correzioni** (le traduzioni esistenti non vengono mai toccate dal flusso normale): correggi nel CSV la colonna `proposta` e metti `s` in `applica` per le righe da scrivere (per i refusi noti è già così), poi **Applica correzioni del controllo a un Excel…** (CLI: `--applica-controllo file.csv --out nuovo.xlsx`). Scrive in una copia, solo nelle celle che contengono ancora il testo controllato, e scarta le proposte con segnaposto `<field>` o numero di righe diversi dall'originale. Le correzioni entrano anche nella memoria come "rivisto" (versioni brevi `@hmi` per le righe HMI).

## File
- `memoria_traduzione.json`: testi già tradotti (parte dal progetto Comipont); cresce a ogni uso. Tienila da parte: è il valore del sistema.
- `glossario.txt`: termini fissi (modificabile a mano: `italiano | inglese | tedesco`, `=` invariato).
- `errori_noti.txt`: refusi e falsi amici noti per il controllo delle traduzioni esistenti (modificabile a mano).
- `classi_allarme.json`: nomi e sigle delle classi per lingua.
- `config.json`: impostazioni (anche da pulsante).

## Regole
- Scrive solo nelle celle vuote, uguali all'originale, "Text" o con soli segnaposto: non sovrascrive mai traduzioni esistenti.
- Le etichette già presenti ma errate (es. schermate HMI con testo italiano diverso dall'originale) non vengono toccate: correggile a mano o dall'Excel.
- Testi di sistema Siemens (~260) restano com'erano.
- Riga di comando opzionale: `python tia_cli.py progetto.xlsx --traduci --residuo residuo.csv --out out.xlsx` (vedi `tia_cli.py`).
