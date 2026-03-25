# Esempio Visivo: Tabella Rischi nel Template Word

## 📊 Struttura Completa della Tabella

Questo documento mostra **esattamente** come deve apparire la tabella dei rischi nel file Word `procedura_template.docx`.

---

## Visualizzazione della Tabella nel Template Word

Apri Microsoft Word e crea una tabella con **10 colonne** e **4 righe** (struttura base):

### Riga 1 - Intestazione della Tabella

| ID | Tipo | Descrizione | Attività/Processo | Probabilità | Gravità | Livello di Rischio | Azioni di Trattamento | Responsabile | Scadenza |
|----|------|-------------|-------------------|-------------|---------|-------------------|----------------------|--------------|----------|

**Cosa scrivere**: Testo normale, nessun tag. Questa è l'intestazione fissa della tabella.

---

### Riga 2 - Inizio del Ciclo FOR

| `{%tr for rischio in rischi %}` | | | | | | | | | |
|---------------------------------|-|-|-|-|-|-|-|-|-|

**⚠️ IMPORTANTE**: 
- Il tag `{%tr for rischio in rischi %}` va nella **prima cella** della riga
- Le altre celle possono essere vuote
- Questo tag dice a docxtpl di iniziare a ripetere le righe successive per ogni rischio

---

### Riga 3 - Dati del Rischio (Riga che viene Ripetuta)

Questa è la riga **PIÙ IMPORTANTE**. Ogni cella contiene i tag per i dati:

| Cellula | Contenuto da Inserire |
|---------|----------------------|
| **Colonna 1 (ID)** | `{{ rischio.id }}` |
| **Colonna 2 (Tipo)** | `{{ rischio.tipo }}` |
| **Colonna 3 (Descrizione)** | `{{ rischio.descrizione }}` |
| **Colonna 4 (Attività/Processo)** | `{{ rischio.attivita_processo }}` |
| **Colonna 5 (Probabilità)** | `{{ rischio.probabilita }}` |
| **Colonna 6 (Gravità)** | `{{ rischio.gravita }}` |
| **Colonna 7 (Livello di Rischio)** | `{% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }}` |
| **Colonna 8 (Azioni)** | `{{ rischio.azioni_trattamento }}` |
| **Colonna 9 (Responsabile)** | `{{ rischio.responsabile }}` |
| **Colonna 10 (Scadenza)** | `{{ rischio.scadenza }}` |

**🎨 Nota sulla Colorazione**: 
Nella colonna 7 "Livello di Rischio", il tag `{% cellbg rischio.codice_colore %}` deve essere **ALL'INIZIO** della cella, subito prima di `{{ rischio.livello_rischio }}`.

Esempio di come appare la cella 7 nel Word:
```
{% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }}
```

Quando lo script esegue il rendering:
- Se `codice_colore` = "FF0000" → la cella diventa **ROSSA** 🔴
- Se `codice_colore` = "FFFF00" → la cella diventa **GIALLA** 🟡
- Se `codice_colore` = "00FF00" → la cella diventa **VERDE** 🟢
- Se `codice_colore` = "FFA500" → la cella diventa **ARANCIONE** 🟠

---

### Riga 4 - Fine del Ciclo FOR

| `{%tr endfor %}` | | | | | | | | | |
|------------------|-|-|-|-|-|-|-|-|-|

**⚠️ IMPORTANTE**: 
- Il tag `{%tr endfor %}` va nella **prima cella** della riga
- Questo tag dice a docxtpl di finire il ciclo di ripetizione

---

## 📋 Esempio Completo Visivo

Ecco come appare l'intera struttura nel Word (con bordi della tabella visibili):

```
┌──────────────┬──────────────┬──────────────────┬─────────────────┬─────────────┬──────────┬──────────────────┬───────────────────┬──────────────┬──────────────┐
│      ID      │     Tipo     │  Descrizione     │Attività/Processo│ Probabilità │ Gravità  │ Livello Rischio  │ Azioni Trattam.   │ Responsabile │   Scadenza   │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│{%tr for risch│              │                  │                 │             │          │                  │                   │              │              │
│io in rischi %│              │                  │                 │             │          │                  │                   │              │              │
│}             │              │                  │                 │             │          │                  │                   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│{{ rischio.id │{{ rischio.ti │{{ rischio.descri │{{ rischio.attiv │{{ rischio.pr│{{ rischio. │{% cellbg rischio│{{ rischio.azioni_ │{{ rischio.re │{{ rischio.sca│
│}}            │po }}         │zione }}          │ita_processo }}  │obabilita }} │gravita }}│.codice_colore %}│ trattamento }}  │sponsabile }}│denza }}      │
│              │              │                  │                 │             │          │{{ rischio.livello│                   │              │              │
│              │              │                  │                 │             │          │_rischio }}       │                   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│{%tr endfor %}│              │                  │                 │             │          │                  │                   │              │              │
└──────────────┴──────────────┴──────────────────┴─────────────────┴─────────────┴──────────┴──────────────────┴───────────────────┴──────────────┴──────────────┘
```

---

## 🎯 Risultato Finale Atteso

Dopo l'esecuzione dello script, la tabella apparirà così (con 5 rischi di esempio):

```
┌──────────────┬──────────────┬──────────────────┬─────────────────┬─────────────┬──────────┬──────────────────┬───────────────────┬──────────────┬──────────────┐
│      ID      │     Tipo     │  Descrizione     │Attività/Processo│ Probabilità │ Gravità  │ Livello Rischio  │ Azioni Trattam.   │ Responsabile │   Scadenza   │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│    R001      │   Rischio    │ Malfunzionamento │   Produzione    │      3      │    4     │    ALTO          │ Manutenzione prev │ Resp. Manut. │ 31/12/2024   │
│              │              │ macchinari       │                 │             │          │  (sfondo         │ programmata...    │              │              │
│              │              │ di produzione    │                 │             │          │   arancione)     │                   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│    R002      │   Rischio    │ Esposizione a    │ Stoccaggio      │      2      │    5     │   ESTREMO        │ DPI obbligatori,  │ Resp. HSE    │ 30/06/2024   │
│              │              │ sostanze chimiche│ materiali       │             │          │  (sfondo rosso)  │ ventilazione...   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│    R003      │   Rischio    │ Ritardo fornitori│Approvvigionam.  │      4      │    3     │    ALTO          │ Qualifica fornit. │ Resp. Acq.   │ 28/02/2024   │
│              │              │ materiali critici│                 │             │          │  (sfondo         │ alternativi...    │              │              │
│              │              │                  │                 │             │          │   arancione)     │                   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│    O001      │ Opportunità  │ Nuova tecnologia │ Gestione        │      4      │    3     │    ALTO          │ Valutazione       │ Resp. Energia│ 30/09/2024   │
│              │              │ per riduzione    │ energetica      │             │          │  (sfondo         │ investimento...   │              │              │
│              │              │ consumi energet. │                 │             │          │   arancione)     │                   │              │              │
├──────────────┼──────────────┼──────────────────┼─────────────────┼─────────────┼──────────┼──────────────────┼───────────────────┼──────────────┼──────────────┤
│    O002      │ Opportunità  │ Espansione       │ Commerciale     │      3      │    4     │    ALTO          │ Studio di mercato │ Dir. Comm.   │ 31/12/2024   │
│              │              │ mercato estero   │                 │             │          │  (sfondo         │ partnership...    │              │              │
│              │              │                  │                 │             │          │   arancione)     │                   │              │              │
└──────────────┴──────────────┴──────────────────┴─────────────────┴─────────────┴──────────┴──────────────────┴───────────────────┴──────────────┴──────────────┘
```

**Nota**: I colori di sfondo nella colonna "Livello di Rischio" saranno:
- 🟢 **Verde** per rischio Basso (punteggio ≤ 4)
- 🟡 **Giallo** per rischio Medio (punteggio 5-9)
- 🟠 **Arancione** per rischio Alto (punteggio 10-16)
- 🔴 **Rosso** per rischio Estremo (punteggio > 16)

---

## ❌ Errori Comuni da Evitare

### Errore 1: Usare `{% for %}` invece di `{%tr for %}`

❌ **SBAGLIATO**:
```
{% for rischio in rischi %}
{{ rischio.id }} | {{ rischio.descrizione }} | ...
{% endfor %}
```

✅ **CORRETTO**:
```
{%tr for rischio in rischi %}
{{ rischio.id }} | {{ rischio.descrizione }} | ...
{%tr endfor %}
```

### Errore 2: Mettere i tag su righe diverse dalla tabella

❌ **SBAGLIATO**:
```
[Tabella]
{%tr for rischio in rischi %}
[Fine tabella]
```

✅ **CORRETTO**:
```
[Riga 1: Intestazione]
[Riga 2: {%tr for rischio in rischi %}]
[Riga 3: {{ rischio.id }} | ...]
[Riga 4: {%tr endfor %}]
```

### Errore 3: Posizionare male il tag cellbg

❌ **SBAGLIATO**:
```
{{ rischio.livello_rischio }}{% cellbg rischio.codice_colore %}
```

✅ **CORRETTO**:
```
{% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }}
```

### Errore 4: Aggiungere il simbolo # ai codici colore

❌ **SBAGLIATO**:
```
codice_colore = "#FF0000"
```

✅ **CORRETTO**:
```
codice_colore = "FF0000"
```

---

## 🔧 Suggerimenti per la Creazione in Word

### Suggerimento 1: Usa "Mostra/Nascondi" ¶

In Word, attiva il pulsante **¶** (Mostra/Nascondi caratteri non stampabili) per vedere esattamente dove sono i tag.

### Suggerimento 2: Allarga le celle temporaneamente

Mentre inserisci i tag, allarga le celle per vedere tutto il testo. Potrai ridimensionare dopo.

### Suggerimento 3: Copia e incolla con cura

Quando copi i tag da questo documento:
1. Copia solo il testo del tag (es. `{%tr for rischio in rischi %}`)
2. Incella nella cella appropriata in Word
3. Verifica che non ci siano spazi extra

### Suggerimento 4: Testa con pochi dati

Prima di creare la procedura completa, testa con un context minimale:
```python
context = {
    "nome_azienda": "Test S.r.l.",
    "rischi": [
        {"id": "R001", "descrizione": "Test", "probabilita": 3, "gravita": 3, ...}
    ]
}
```

---

## 📸 Screenshot Concettuale

Immagina la tabella in Word così (vista con griglia visibile):

```
╔══════════════════════════════════════════════════════════════════════════════════════════╗
║  ID  │ Tipo │ Descrizione │ ... │ Livello Rischio        │ ... │ Scadenza                ║
╠══════════════════════════════════════════════════════════════════════════════════════════╣
║ {%tr for rischio in rischi %}                                                              ║
╠══════════════════════════════════════════════════════════════════════════════════════════╣
║ {{r.id}}│{{r.tipo}}│{{r.desc}}│...│{% cellbg r.colore %}{{r.livello}}│...│{{r.scad}}   ║
╠══════════════════════════════════════════════════════════════════════════════════════════╣
║ {%tr endfor %}                                                                             ║
╚══════════════════════════════════════════════════════════════════════════════════════════╝
```

---

## ✅ Checklist Finale

Prima di salvare il template, verifica:

- [ ] La tabella ha 10 colonne
- [ ] La riga con `{%tr for rischio in rischi %}` è una riga separata
- [ ] La riga con i dati `{{ rischio.xxx }}` è una riga separata
- [ ] La riga con `{%tr endfor %}` è una riga separata
- [ ] Il tag `{% cellbg rischio.codice_colore %}` è ALL'INIZIO della cella "Livello Rischio"
- [ ] Non ci sono simboli `#` nei codici colore
- [ ] Tutti i tag hanno due parentesi graffe `{{ }}` o `{% %}`
- [ ] Il file è salvato come `.docx`

---

**Buon lavoro!** 🚀

Se hai seguito tutte le istruzioni, il tuo template genererà una tabella professionale con:
- ✅ Righe dinamiche per ogni rischio/opportunità
- ✅ Colorazione automatica dello sfondo in base al livello di rischio
- ✅ Tutti i dati formattati correttamente
