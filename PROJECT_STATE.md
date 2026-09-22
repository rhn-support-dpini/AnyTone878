# PROJECT_STATE — AT-DvP

Ultimo aggiornamento: 2026-09-22.

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
| Zone all'avvio | 2 = **DigiAL** su entrambi i lati (`StartZone1/2=2`) |
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

1. Analog — 2. **DigiAL** (IR1UGF + IK1HJT + IR1UIZ + IR1ZVJ + IR1UIW + IR1UIV + IR1ZZX) — 3. AIB — 4. Emergenza — 5. pmr — 6. lpd — 7. CRI — 8. D2ALP — 9. Marini — 10. vhf-uhf

### IR1UGF (Ponzone, BrandMeister)

| Canale | Slot | Contatto/TG |
|--------|------|-------------|
| Worldwide, Europe, Italia | **1** | TG 222 |
| Loca1, Cluster | **1** | TG 222 |
| Discon, Loca2, Parrot, APRS | **2** | Parrot 222997 / APRS 222999 |

- **IR1UGF-Italia (#83):** D-APRS attivo (Report Channel 1 → IR1UGF APRS #78)
- **APRS Auto TX Interval:** 60 s
- Prima attivazione APRS sulla radio: SMS `@SSID 7` in privato a 222999

### IK1HJT (Arquata, BrandMeister)

| Parametro | Valore |
|-----------|--------|
| Downlink (RX) | 430.300 MHz |
| Uplink (TX) | 437.900 MHz |
| Offset | +7.6 MHz |
| Color code | 1 |

Stack completo (#63, #70–#77): Worldwide, Europe, Italia, Discon, Loca1, Cluster, Loca2, Parrot, APRS — stessa struttura IR1UGF.

| Canale | Slot | Note |
|--------|------|------|
| Worldwide, Europe, Italia | **1** | TG 222 |
| Loca1, Cluster | **1** | TG 222 |
| Discon, Loca2, Parrot (#71), APRS (#70) | **2** | Parrot 222997 / APRS 222999 |

- **IK1HJT-Italia (#75):** D-APRS attivo (Report Channel 1 → IK1HJT APRS #70)
- Rimosso vecchio canale singolo `IK1HJT Arquata` dalla zona D2ALP

### IR1UIZ (Giarole, BrandMeister)

| Parametro | Valore |
|-----------|--------|
| Downlink (RX) | 431.43750 MHz |
| Uplink (TX) | 433.03750 MHz |
| Offset | +1.6 MHz |
| Color code | 5 |

Stack (#105–#116 + legacy TG #107–#109): TG88/TG222/APRS legacy + Worldwide, Europe, Italia, Discon, Loca1, Cluster, Loca2, Parrot, APRS — tutti in DigiAL/scan **AlDigiScan**. Nomi **`IR1UIZ-*`** (prefisso `AL-` rimosso).

| Canale | Slot | Note |
|--------|------|------|
| Worldwide, Europe, Italia | **1** | TG 222 |
| Loca1, Cluster | **1** | TG 222 |
| Discon, Loca2, Parrot (#106), APRS (#105) | **2** | Parrot 222997 / APRS 222999 |

- **IR1UIZ-Italia (#113):** D-APRS attivo (Report Channel 1 → IR1UIZ APRS #105)
- **APRS.CSV:** `channel3=105`/`slot3=2` (IR1UIZ APRS), TG 222999
- **Scan list AlDigiScan / Zona DigiAL:** vedi riepilogo sotto (66 canali totali)
- Vecchie copie incomplete `AL-IR1UIZ-*` in memoria vhf-uhf (#431–#437) svuotate (duplicati di nome)

### IR1ZVJ / IR1UIW / IR1UIV / IR1ZZX (vhf-uhf, CC 1)

Stack 9 canali ciascuno (#87–#134, ordine alfabetico nel blocco DigiAL), inseriti in memoria dopo IR1UIZ con insert+shift; parametri dalle righe `AL-*` preesistenti in vhf-uhf (poi svuotate).

| Ponte | RX | TX | Memoria (post-ordine alfabetico) |
|-------|-----|-----|---------|
| IR1UIV | 431.50000 | 433.10000 | #87–#95 |
| IR1UIW | 431.28750 | 432.88750 | #96–#104 |
| IR1ZVJ | 430.28750 | 431.88750 | #117–#125 |
| IR1ZZX | 431.52500 | 433.12500 | #126–#134 |

Stessa struttura slot/contatti di IR1UGF; D-APRS su `*-Italia`; scan list **AlDigiScan** (+ **AlDigiScan2** per limite CPS 50 canali).

### DigiAL — riepilogo (66 canali)

Blocco memoria **#69–#133** ordinato alfabeticamente per nome canale (riga file Channel 70–134); **#63** (`IK1HJT-Worldw`) e **#134** (`IR1ZZX APRS`) fuori dal blocco ordinato.

| Ponte | Canali in zona/scan |
|-------|---------------------|
| IR1UGF | 9 (#78–#86) — nomi `IR1UGF-*` |
| IK1HJT | 9 (#63, #70–#77) — nomi `IK1HJT-*` |
| IR1UIZ | 12 (#105–#116, legacy TG #107–#109) — nomi `IR1UIZ-*` |
| IR1UIV / IR1UIW / IR1ZVJ / IR1ZZX | 9 ciascuno (#87–#95, #96–#104, #117–#125, #126–#134) |

- **Boot zona:** A = `IR1UGF-Cluste`, B = `IR1UGF APRS` (invariato)
- **APRS.CSV:** `channel1=78`, `channel2=70`, `channel3=105` (IR1UGF / IK1HJT / IR1UIZ APRS)
- **Scan list DigiAL:** `AlDigiScan` (#2, 42 canali operativi) + `DigiService` (#3, 12 Parrot/APRS dei 6 ponti IR1UGF/IK1HJT/IR1ZVJ/IR1UIW/IR1UIV/IR1ZZX). IR1UIZ fuori da entrambe.

### Scan list **AlAnaScan** (#2)

Tutti i **12 canali analogici `AL-*`** (prefisso provinciale AL), ordinati alfabeticamente. Canali con `Scan List = AlAnaScan`.

**Limite CPS:** max **50 canali** per scan list (oltre → *Import Error*).

---

## Modifiche applicate (cronologia)

### Sessione 2026-09-22

1. **Fix inserimento IK1HJT:** i canali `PC Ponte` (#90), `PC Diretta` (#91) e `AIB Diretta 1` (#92) erano stati sovrascritti dallo stack IK1HJT (#82–#89); ripristinati con inserimento a #90 e shift delle righe successive (688 slot totali invariati). Zona/scan AIB già corretti per nome.
2. **IR1UIZ stack DigiAL:** inseriti 9 canali (#90–#98) dopo IK1HJT APRS con insert+shift (AIB da #99); parametri da righe IR1UIZ esistenti (431.43750/433.03750, CC 5); zona DigiAL, scan Digital e APRS channel3 aggiornati.
3. **IR1UIZ completo in DigiAL/scan:** aggiunti anche i 3 canali legacy `IR1UIZ TG 88/222/222999` (#70–#72) con scan list Digital; totale IR1UIZ in zona/scan: 12 canali.
4. **Rinomina AL- → rimosso** su canali `IR1UGF-*` e `IK1HJT-*` (Channel, Zone, ScanList e CSV correlati).
5. **Nuovi stack DigiAL:** IR1ZVJ, IR1UIW, IR1UIV, IR1ZZX (36 canali #99–#134) inseriti con shift; legacy `AL-IR1ZVJ/UIW/UIV/ZZX-*` in vhf-uhf svuotati; DigiAL/scan Digital a 66 canali.
6. **Rinomina `AL-IR1UIZ-*` → `IR1UIZ-*`** in Channel, Zone, ScanList e CSV correlati.
7. **Ordine alfabetico** canali DigiAL in memoria (#69–#133, righe file 70–134); aggiornati riferimenti numerici in `APRS.CSV`.
8. **Scan list AlAnaScan / AlDigiScan / DigiService:** analogici `AL-*` (12); digitali 6 ponti in `AlDigiScan` (42) e Parrot/APRS in `DigiService` (12).

### Sessione 2026-09-21

1. **IK1HJT Arquata:** stack 9 canali (430.300/437.900, CC 1), scan Digital, D-APRS su Italia + APRS #89.
2. **IR1UGF Parrot:** `Busy Lock/TX Permit=Always` (TX abilitata).
3. **Friend/contact Enzo IK1EVM:** DMR ID 2221707 in `TalkGroups.CSV` + `Call Alert=Online Alert` in `DigitalContactList.CSV`.
4. **Scan list Digital:** solo 9 canali IR1UGF; rimossi D2ALP e IR1UIZ.
5. **Scan list Altri:** eliminata (canali OS4/PS2/IR3 restano in memoria, fuori zona).
6. **Scan list CRI:** tutti i 24 canali CRI con `Scan List=CRI`.
7. **Canale analogico APRS** (#1, 144.8 MHz): rimosso (slot vuoto).
8. **Header CSV:** corretti `OptionalSetting.CSV` e `APRS.CSV` (campi concatenati + riga dati APRS spezzata).
9. **Fix import CPS:** `APRS.CSV` riga dati unificata (232 colonne); `OptionalSetting.CSV` aggiunto valore mancante `SateAosLimit=0`.
10. **Slot DMR IR1UGF/IK1HJT:** slot **1** = Worldwide/Europe/Italia/Loca1/Cluster; slot **2** = Discon/Loca2/Parrot/APRS (`APRS.CSV` slot1/slot2=2).
11. **Zona Digital → DigiAL:** rinominata; aggiunti 9 canali IK1HJT (18 canali totali in zona #2).

### Sessione 2026-09-20

1. Audit generale: 621 canali, 10 zone, 8 scan list — nessun errore strutturale grave.
2. **CRI** (24 canali): potenza High, CTCSS decode 156.7.
3. **Scan Analog:** aggiunta Diretta; **Scan CRI:** rimossi RRM ch 8/16.
4. **RIP G ANAL** → AIB Rpt Ana G, potenza High.
5. **Roaming:** eliminati 4 canali US 410 MHz; Roam Zone 1 vuota.
6. Riordino zone e memoria canali (#1–#688, gruppi con 5 righe vuote).
7. **IR1UGF:** canali Parrot + APRS, Talk Group Parrot 222997, GPS/APRS in OptionalSetting.
8. Parrot su slot 2; intervallo invio posizione **60 s**.
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
