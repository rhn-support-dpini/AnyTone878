# PROJECT_STATE — AT-DvP

Ultimo aggiornamento: 2026-09-20.

Documento di handoff per sessioni successive: radio, struttura repo, modifiche recenti, convenzioni.

---

## Panoramica del progetto

**AT-DvP** — configurazione e automazione per **AnyTone AT-D878UVII Plus** (GPS + BT), callsign **IU1FLA**, DMR ID **2221851**.

| Componente | Ruolo |
|------------|--------|
| `CurrentConfiguration/` | Export CSV dal CPS AnyTone (fonte di verità per modifiche) |
| `Script1_BuildExcel.py` | CSV → Excel `Master_Codeplug` (revisione visiva) |
| `Script2_ExportCSV.py` | Excel → CSV per re-import nel CPS |
| `codeplug_utils.py` | Utilità condivise (CSV, frequenze, log) |
| `start/` | Snapshot iniziali (2026-09-08) per confronto |
| `requirements.txt` | Dipendenze Python (`openpyxl`) |

**Workflow CPS:** modificare i CSV in `CurrentConfiguration/` → re-importare nel CPS → scrivere sulla radio.  
**Workflow Excel (opzionale):** `python3 Script1_BuildExcel.py .` → edit Excel → `python3 Script2_ExportCSV.py .`

**Repo:** `https://github.com/rhn-support-dpini/AT-DvP.git` — branch `main`.

---

## Radio — impostazioni principali

| Parametro | Valore |
|-----------|--------|
| Modello | AT-D878UVII Plus |
| Callsign | IU1FLA |
| DMR ID | 2221851 |
| Zone all'avvio | 1 = Analog, 2 = Digital (`StartZone1/2`) |
| Canale all'avvio | Ripristina ultimo usato (`StartChUse=1`) |
| GPS | On, modalità APRS, beacon posizione |

### Zone (ordine 1–10)

1. Analog — 2. Digital — 3. AIB — 4. Emergenza — 5. pmr — 6. lpd — 7. CRI — 8. D2ALP — 9. Marini — 10. vhf-uhf

### IR1UGF (Ponzone, BrandMeister)

| Canale | RX/TX MHz | CC | Slot | Contatto/TG |
|--------|-----------|-----|------|-------------|
| IR1UGF Parrot | 431.225 / 432.825 | 1 | **2** | Private 222997 |
| IR1UGF APRS | 431.225 / 432.825 | 1 | 2 | Private 222999 |

- **AL-IR1UGF-Italia:** D-APRS attivo (Report Channel 1 → canale APRS)
- **APRS Auto TX Interval:** 60 s
- Prima attivazione APRS sulla radio: SMS `@SSID 7` in privato a 222999

---

## Modifiche applicate (cronologia)

### Sessione 2026-09-20

1. Audit generale: 621 canali, 10 zone, 8 scan list — nessun errore strutturale grave.
2. **CRI** (24 canali): potenza High, CTCSS decode 156.7.
3. **Scan Analog:** aggiunta Diretta; **Scan CRI:** rimossi RRM ch 8/16.
4. **RIP G ANAL** → AIB Rpt Ana G, potenza High.
5. **Roaming:** eliminati 4 canali US 410 MHz; Roam Zone 1 vuota.
6. Riordino zone e memoria canali (#1–#688, gruppi con 5 righe vuote).
7. **IR1UGF:** canali Parrot + APRS, Talk Group Parrot 222997, GPS/APRS in OptionalSetting.
8. Parrot su **slot 2**; intervallo invio posizione **60 s**.

### Baseline

- Snapshot `start/2026-09-08-*.csv` — export prima delle modifiche strutturali.

---

## Convenzioni

- **Directory attiva:** solo `CurrentConfiguration/` per edit CPS (non file CSV sparsi in root).
- **Commit:** solo su richiesta esplicita dell'utente.
- **Stile commit:** `feat(config): …`, `fix(channel): …`, `docs: …`.
- **Whitelist `.gitignore`:** tracciati solo script, CSV CPS, docs e regole Cursor; log e artefatti generati esclusi.

---

## Comandi utili

```bash
cd /home/dpini/AnyTone878   # workspace locale (progetto AT-DvP)

python3 -m pip install -r requirements.txt
python3 Script1_BuildExcel.py .
python3 Script2_ExportCSV.py .
```

---

## Prossimi passi suggeriti

1. Re-importare `CurrentConfiguration/` nel CPS e scrivere sulla radio dopo ogni modifica significativa.
2. Rigenerare `AnyTone.rdt` dal CPS se serve backup binario (non versionato in git).
3. Verificare su BrandMeister/APRS.fi ricezione beacon IR1UGF dopo attivazione GPS.

---

*Aggiornare questo file alla fine di ogni sessione significativa.*
