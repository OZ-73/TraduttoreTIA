# Traduttore TIA

Traduce i testi di un progetto TIA Portal (Excel esportato da TIA) con memoria di traduzione, glossario e modello locale (LM Studio).
Vedi `LEGGIMI.md` per l'uso. L'exe per Windows viene generato da GitHub Actions (scheda Actions → ultimo run → Artifacts).

La memoria `memoria_traduzione.json` non è nel repository (dati dei progetti): va tenuta accanto all'exe.

## Stato (10/11/2026)
- App tkinter + exe (GitHub Actions → Artifacts "TraduttoreTIA-windows"); memoria `memoria_traduzione.json` fuori dal repository.
- Opzioni: "solo testi di allarme" (categorie e percorsi in `config.json`), "non tradurre passi/transizioni GRAPH" (riconosciuti dal percorso `SYSTEM_StepNames`/`SYSTEM_TransitionNames`).
- Glossario `glossario.txt` aggiornato con i termini del progetto OMB (rivettatrici, foratrici, cassonetti, ecc.).
