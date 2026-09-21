# PROJECT_STATE — AT-DvP

Ultimo aggiornamento: 2026-09-21.

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
| Zone all'avvio | 2 = Digital su entrambi i lati (`StartZone1/2=2`) |
| Canale all'avvio | Fisso: Cluster (#6) + APRS (#9) (`StartChUse=0`) |
| Display | Schermo diviso (`SubMode=1`, `DiviDisEn=1`); backlight **30 s** (`AutoBKLightTime=6`) |
| GPS | On, modalità APRS, beacon posizione |

### Display e tasti programmabili (non in CSV)

L'export CSV **non include** le assegnazioni tasti (`P1 Long`, `P2 Long`, ecc.). Vanno reimpostate **manualmente nel CPS** dopo ogni import CSV, prima di scrivere sulla radio:

**CPS → Public/Common Setting → Optional Setting → Key Function**

| Tasto | Pressione | Funzione CPS |
|-------|-----------|--------------|
| **P1** | Long | **Main CH Switch** (alterna canale A/B sul display) |
| **P2** | Long | **Sub CH Hide** (visualizzazione 1 o 2 canali) |

Backlight 30 s: `OptionalSetting.CSV` → `AutoBKLightTime=6` (già impostato).

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

### IK1HJT (Arquata, BrandMeister)

| Parametro | Valore |
|-----------|--------|
| Downlink (RX) | 430.300 MHz |
| Uplink (TX) | 437.900 MHz |
| Offset | +7.6 MHz |
| Color code | 1 |

Stack completo (#63, #82–#89): Worldwide, Europe, Italia, Discon, Loca1, Cluster, Loca2, Parrot, APRS — stessa struttura IR1UGF.

| Canale | Slot | Note |
|--------|------|------|
| IK1HJT Parrot (#88) | 2 | Private 222997, TX Always |
| IK1HJT APRS (#89) | 2 | Private 222999, D-APRS |

- **AL-IK1HJT-Italia (#83):** D-APRS attivo (Report Channel 1 → IK1HJT APRS)
- **APRS.CSV:** `channel2=89`, `slot2=2`, TG 222999 (secondo ripetitore)
- **Scan list Digital:** 18 canali (9 IR1UGF + 9 IK1HJT)
- **Zona Digital:** resta solo IR1UGF (boot invariato); IK1HJT in memoria + scan
- Rimosso vecchio canale singolo `IK1HJT Arquata` dalla zona D2ALP

---

## Modifiche applicate (cronologia)

### Sessione 2026-09-21

1. **IK1HJT Arquata:** stack 9 canali (430.300/437.900, CC 1), scan Digital, D-APRS su Italia + APRS #89.
2. **IR1UGF Parrot:** `Busy Lock/TX Permit=Always` (TX abilitata).
3. **Friend/contact Enzo IK1EVM:** DMR ID 2221707 in `TalkGroups.CSV` + `Call Alert=Online Alert` in `DigitalContactList.CSV`.
4. **Scan list Digital:** solo 9 canali IR1UGF; rimossi D2ALP e IR1UIZ.
5. **Scan list Altri:** eliminata (canali OS4/PS2/IR3 restano in memoria, fuori zona).
6. **Scan list CRI:** tutti i 24 canali CRI con `Scan List=CRI`.
7. **Canale analogico APRS** (#1, 144.8 MHz): rimosso (slot vuoto).
8. **Header CSV:** corretti `OptionalSetting.CSV` e `APRS.CSV` (campi concatenati + riga dati APRS spezzata).

### Sessione 2026-09-20

1. Audit generale: 621 canali, 10 zone, 8 scan list — nessun errore strutturale grave.
2. **CRI** (24 canali): potenza High, CTCSS decode 156.7.
3. **Scan Analog:** aggiunta Diretta; **Scan CRI:** rimossi RRM ch 8/16.
4. **RIP G ANAL** → AIB Rpt Ana G, potenza High.
5. **Roaming:** eliminati 4 canali US 410 MHz; Roam Zone 1 vuota.
6. Riordino zone e memoria canali (#1–#688, gruppi con 5 righe vuote).
7. **IR1UGF:** canali Parrot + APRS, Talk Group Parrot 222997, GPS/APRS in OptionalSetting.
8. Parrot su **slot 2**; intervallo invio posizione **60 s**.
9. Zona Digital: solo canali IR1UGF; avvio su Cluster + APRS; schermo diviso.
10. **Backlight 30 s** (`AutoBKLightTime=6`); tasti P1/P2 long documentati (solo CPS).

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

1. Re-importare `CurrentConfiguration/` nel CPS.
2. **Reimpostare Key Function** (P1 Long, P2 Long) nel CPS — non esportate nei CSV.
3. Scrivere il codeplug sulla radio.
4. Rigenerare `AnyTone.rdt` dal CPS come backup binario (non versionato in git).

---

*Aggiornare questo file alla fine di ogni sessione significativa.*
