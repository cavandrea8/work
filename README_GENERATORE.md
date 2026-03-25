# Generatore Procedura Gestione Rischi e Opportunità ISO 9001/14001/45001

## 📋 Descrizione

Script Python professionale per la generazione automatica della **Procedura per la Gestione dei Rischi e delle Opportunità** integrata per le norme:
- **ISO 9001:2015** (Qualità)
- **ISO 14001:2015** (Ambiente)
- **ISO 45001:2018** (Sicurezza sul Lavoro)

Lo script utilizza la libreria **docxtpl** per generare documenti Word partendo da template con tag Jinja2.

---

## 🚀 Novità Versione 2.0

- ✅ **Sintassi corretta docxtpl** per tabelle dinamiche (`{%tr for %}` ... `{%tr endfor %}`)
- ✅ **Colorazione automatica celle** con `{% cellbg codice_colore %}`
- ✅ **Supporto InlineImage** per logo aziendale e immagini nei rischi
- ✅ **Export PDF opzionale** tramite docx2pdf
- ✅ **Logging avanzato** su file e console
- ✅ **Gestione migliorata delle eccezioni**
- ✅ **Validazione completa** del context e dei file

---

## 📦 Installazione

### Requisiti

- Python 3.7 o superiore
- Microsoft Word o LibreOffice (per export PDF)

### Installa le dipendenze

```bash
pip install docxtpl
pip install docx2pdf  # Opzionale, solo per export PDF
```

---

## 📁 File del Progetto

| File | Descrizione |
|------|-------------|
| `generatore_procedura_rischi_v2.py` | Script Python principale |
| `procedura_template.docx` | Template Word (DA CREARE) |
| `ISTRUZIONI_TEMPLATE.md` | Istruzioni dettagliate per creare il template |
| `ESEMPIO_TABELLA_RISCHI.md` | Esempio visivo della tabella rischi |
| `README_GENERATORE.md` | Questo file |

---

## 🔧 Utilizzo

### Utilizzo Base

```bash
python generatore_procedura_rischi_v2.py
```

Lo script:
1. Carica il template `procedura_template.docx`
2. Valida i dati di esempio
3. Genera il documento Word
4. (Opzionale) Converte in PDF se richiesto

### Utilizzo Personalizzato

```python
from generatore_procedura_rischi_v2 import GeneratoreProceduraRischi

# Crea istanza del generatore
generatore = GeneratoreProceduraRischi("procedura_template.docx")

# Definisci i tuoi dati
context = {
    "nome_azienda": "La Tua Azienda S.r.l.",
    "responsabile": "Ing. Mario Verdi",
    "revisione": "01",
    "data_emissione": "01/02/2024",
    "rischi": [
        {
            "id": "R001",
            "descrizione": "Rischio esempio",
            "probabilita": 3,
            "gravita": 4,
            # ... altri campi
        }
    ]
}

# Valida
errori = generatore.valida_context(context)
if errori:
    print("Errori:", errori)
else:
    # Genera
    generatore.genera_documento(context, output_path="Procedura_Personalizzata.docx")
```

---

## 🎯 Come Creare il Template Word

### Passo 1: Leggi le Istruzioni

Apri il file [`ISTRUZIONI_TEMPLATE.md`](./ISTRUZIONI_TEMPLATE.md) per le istruzioni complete passo-passo.

### Passo 2: Guarda l'Esempio della Tabella

Apri [`ESEMPIO_TABELLA_RISCHI.md`](./ESEMPIO_TABELLA_RISCHI.md) per vedere esattamente come deve essere strutturata la tabella dei rischi.

### Passo 3: Crea il Template

1. Apri Microsoft Word
2. Crea un nuovo documento
3. Inserisci le sezioni come descritto nelle istruzioni
4. **Importante**: Usa la sintassi corretta per le tabelle dinamiche:
   - `{%tr for rischio in rischi %}` per iniziare il ciclo
   - `{%tr endfor %}` per terminare il ciclo
   - `{% cellbg rischio.codice_colore %}` per colorare le celle

### Passo 4: Salva

Salva il file come `procedura_template.docx` nella stessa cartella dello script.

---

## 📊 Struttura del Context

Il context è un dizionario Python con tutti i dati per il template:

### Dati Obbligatori

```python
{
    "nome_azienda": "Nome Azienda",
    "responsabile": "Nome Responsabile",
    "revisione": "01",
    "data_emissione": "GG/MM/AAAA"
}
```

### Dati Rischi/Oppurtunità

Ogni rischio/opportunità deve avere:

```python
{
    "id": "R001",
    "descrizione": "Descrizione del rischio",
    "attivita_processo": "Processo aziendale",
    "tipo": "Rischio",  # o "Opportunità"
    "probabilita": 3,   # 1-5
    "gravita": 4,       # 1-5
    "azioni_trattamento": "Azioni da intraprendere",
    "responsabile": "Responsabile azione",
    "scadenza": "GG/MM/AAAA"
}
```

**Nota**: Il livello di rischio viene calcolato automaticamente dallo script!

### Logo Aziendale (Opzionale)

```python
{
    "logo_path": "percorso/logo.png"
}
```

---

## 🎨 Colorazione Automatica dei Rischi

Lo script calcola automaticamente il livello di rischio e applica i colori:

| Punteggio (P×G) | Livello | Colore Sfondo | Codice |
|-----------------|---------|---------------|--------|
| ≤ 4 | Basso | 🟢 Verde | `00FF00` |
| 5-9 | Medio | 🟡 Giallo | `FFFF00` |
| 10-16 | Alto | 🟠 Arancione | `FFA500` |
| > 16 | Estremo | 🔴 Rosso | `FF0000` |

---

## 📝 Esempio di Output

Dopo l'esecuzione, otterrai un documento Word con:

- ✅ Intestazione aziendale con logo (se fornito)
- ✅ Tutte le sezioni della procedura ISO
- ✅ Tabelle con scale di probabilità e gravità
- ✅ **Tabella dinamica dei rischi** con righe generate automaticamente
- ✅ **Celle colorate** in base al livello di rischio
- ✅ Flusso operativo e frequenze di monitoraggio
- ✅ Record correlati

---

## ⚠️ Errori Comuni e Soluzioni

### Errore: "Template non trovato"

**Soluzione**: Assicurati che `procedura_template.docx` esista nella stessa cartella dello script.

### Errore: Tabella con una sola riga invece di multiple

**Causa**: Hai usato `{% for %}` invece di `{%tr for %}`

**Soluzione**: Nel template Word, usa:
```
{%tr for rischio in rischi %}
...
{%tr endfor %}
```

### Errore: Colori non applicati alle celle

**Causa 1**: Il tag `{% cellbg %}` non è all'inizio della cella

**Soluzione**: Usa:
```
{% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }}
```

**Causa 2**: Codice colore con simbolo `#`

**Soluzione**: Usa `FF0000`, non `#FF0000`

### Errore: Export PDF fallito

**Causa**: Mancanza di Microsoft Word o LibreOffice

**Soluzione**: 
- Installa Microsoft Word, oppure
- Usa solo l'output DOCX, oppure
- Converti manualmente con un altro strumento

---

## 📄 Log e Debug

Lo script crea un file di log `generatore_procedura.log` con:
- Timestamp di ogni operazione
- Errori e warning
- Dettagli sul rendering

Per visualizzare i log in tempo reale:
```bash
python generatore_procedura_rischi_v2.py
tail -f generatore_procedura.log  # Su Linux/Mac
```

---

## 🔐 Validazione dei Dati

Prima di generare il documento, lo script valida:

1. ✅ Campi obbligatori presenti
2. ✅ Probabilità e gravità tra 1 e 5
3. ✅ Descrizioni rischi presenti
4. ✅ File logo esistenti (se specificati)
5. ✅ File immagini rischi esistenti (se specificati)

Se ci sono errori, lo script si ferma e mostra un report dettagliato.

---

## 📚 Riferimenti Normativi

La procedura generata copre i requisiti di:

- **ISO 9001:2015** - Paragrafo 6.1 (Azioni per affrontare rischi e opportunità)
- **ISO 14001:2015** - Paragrafo 6.1 (Azioni per affrontare rischi e opportunità ambientali)
- **ISO 45001:2018** - Paragrafo 6.1 (Azioni per affrontare rischi e opportunità SSL)
- **D.Lgs. 81/2008** - Valutazione dei Rischi

---

## 💡 Suggerimenti per i Consulenti ISO

1. **Personalizza il template**: Aggiungi il tuo logo e intestazione standard
2. **Crea contest riutilizzabili**: Salva i context per clienti diversi
3. **Mantieni un archivio**: Conserva le procedure generate per ogni cliente
4. **Aggiorna regolarmente**: Tieni traccia delle revisioni nel tempo
5. **Verifica sempre**: Controlla il documento generato prima di approvarlo

---

## 🆘 Supporto

In caso di problemi:

1. Controlla il file di log `generatore_procedura.log`
2. Verifica di aver seguito le istruzioni in `ISTRUZIONI_TEMPLATE.md`
3. Controlla gli esempi in `ESEMPIO_TABELLA_RISCHI.md`
4. Testa con un template minimale prima di creare quello completo

---

## 📄 Licenza

Script sviluppato per uso professionale da consulenti di sistemi di gestione integrati.

---

## 👨‍💻 Autore

Consulente Sistemi di Gestione Integrati  
Versione: 2.0  
Data: 2024

---

## 🔄 Changelog

### v2.0 (2024)
- ✅ Aggiunta sintassi corretta `{%tr for %}` per tabelle dinamiche
- ✅ Supporto colorazione celle con `{% cellbg %}`
- ✅ Supporto InlineImage per logo e immagini
- ✅ Export PDF opzionale
- ✅ Logging su file
- ✅ Migliore gestione eccezioni
- ✅ Validazione avanzata del context

### v1.0 (Precedente)
- Versione iniziale

---

**Buon lavoro!** 🚀

Per iniziare:
1. Leggi `ISTRUZIONI_TEMPLATE.md`
2. Crea il tuo template Word
3. Esegui lo script
4. Genera la tua procedura ISO professionale!
