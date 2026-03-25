# Istruzioni per Creare il Template Word per la Procedura di Gestione Rischi e Opportunità ISO 9001/14001/45001

## 📋 Panoramica

Questo documento fornisce istruzioni **passo-passo** per creare il file `procedura_template.docx` che sarà utilizzato dallo script Python per generare automaticamente la procedura di gestione rischi e opportunità.

Il template utilizza la libreria **docxtpl** che supporta i tag **Jinja2** direttamente in documenti Word.

---

## 🛠️ Prerequisiti

1. **Microsoft Word** (consigliato) o LibreOffice Writer
2. **Python 3.7+** con le librerie installate:
   ```bash
   pip install docxtpl docx2pdf
   ```

---

## 📝 Struttura del Template

Il template deve essere un normale documento Word (.docx) con:

### 1. **Intestazione del Documento**

Nella prima pagina, crea una tabella 2x2 per l'intestazione:

```
┌─────────────────────────────────┬─────────────────────────────────┐
│ {{ logo }}                      │ Codice Documento: {{ codice_documento }} │
│ (se presente)                   │ Revisione: {{ revisione }}               │
│                                 │ Data Emissione: {{ data_emissione }}     │
│                                 │ Pagina: {{ pagina }}                     │
└─────────────────────────────────┴─────────────────────────────────┘
```

**Nota importante**: Il tag `{{ logo }}` mostrerà automaticamente il logo aziendale se fornito nel context. Se non fornito, rimarrà vuoto.

---

### 2. **Titolo del Documento**

Centrato, in grassetto, dimensione 16-18pt:

```
PROCEDURA PER LA GESTIONE DEI RISCHI E DELLE OPPORTUNITÀ
Sistema di Gestione Integrato ISO 9001/14001/45001
```

---

### 3. **Sezione 1 - Scopo e Campo di Applicazione**

```
1. SCOPO E CAMPO DI APPLICAZIONE

La presente procedura definisce le modalità per l'identificazione, la valutazione 
e la gestione dei rischi e delle opportunità relativi al Sistema di Gestione Integrato 
per la Qualità (ISO 9001:2015), l'Ambiente (ISO 14001:2015) e la Sicurezza sul Lavoro 
(ISO 45001:2018).

Si applica a tutti i processi aziendali di {{ nome_azienda }}.
```

---

### 4. **Sezione 2 - Riferimenti Normativi**

```
2. RIFERIMENTI NORMATIVI

{{ riferimenti_normativi_text }}
```

**Nota**: Lo script formatta automaticamente la lista in punti elenco.

---

### 5. **Sezione 3 - Definizioni**

Crea una tabella 2x5 per le definizioni:

```
┌──────────────────────┬─────────────────────────────────────────────┐
│ Termine              │ Definizione                                 │
├──────────────────────┼─────────────────────────────────────────────┤
│ Rischio              │ {{ definizioni.rischio }}                   │
├──────────────────────┼─────────────────────────────────────────────┤
│ Opportunità          │ {{ definizioni.opportunita }}               │
├──────────────────────┼─────────────────────────────────────────────┤
│ Parte Interessata    │ {{ definizioni.parte_interessata }}         │
├──────────────────────┼─────────────────────────────────────────────┤
│ Valutazione Rischio  │ {{ definizioni.valutazione_rischio }}       │
└──────────────────────┴─────────────────────────────────────────────┘
```

---

### 6. **Sezione 4 - Responsabilità e Autorità**

```
4. RESPONSABILITÀ E AUTORITÀ

Responsabile della procedura: {{ responsabile }}
Ruolo: {{ ruolo_responsabile }}

Approvato da: {{ approvato_da }}
Ruolo: {{ ruolo_approvatore }}
Data approvazione: {{ data_approvazione }}
```

---

### 7. **Sezione 5 - Matrice di Valutazione del Rischio**

#### 7.1 Scala di Probabilità

Crea una tabella con 6 righe (1 intestazione + 5 dati):

```
┌───────┬────────────────────┬──────────────────────────────────────┐
│ Valore│ Descrizione        │ Frequenza                            │
├───────┼────────────────────┼──────────────────────────────────────┤
{% for p in scala_probabilita %}
│ {{ p.valore }} │ {{ p.descrizione }} │ {{ p.frequenza }} │
{% endfor %}
└───────┴────────────────────┴──────────────────────────────────────┘
```

**IMPORTANTE**: Nel template Word, la struttura deve essere ESATTAMENTE così:

| Riga 1 (intestazione): | Valore | Descrizione | Frequenza |
|------------------------|--------|-------------|-----------|
| Riga 2: | `{%tr for p in scala_probabilita %}` | | |
| Riga 3: | `{{ p.valore }}` | `{{ p.descrizione }}` | `{{ p.frequenza }}` |
| Riga 4: | `{%tr endfor %}` | | |

#### 7.2 Scala di Gravità

Stessa struttura della probabilità:

```
┌───────┬────────────────────┬──────────────────────────────────────┐
│ Valore│ Descrizione        │ Impatto                              │
├───────┼────────────────────┼──────────────────────────────────────┤
{%tr for g in scala_gravita %}
│ {{ g.valore }} │ {{ g.descrizione }} │ {{ g.impatto }} │
{%tr endfor %}
└───────┴────────────────────┴──────────────────────────────────────┘
```

---

### 8. **Sezione 6 - Elenco Rischi e Opportunità** ⭐

Questa è la sezione **PIÙ IMPORTANTE**. La tabella deve avere la seguente struttura:

#### Intestazione della Tabella

Crea una tabella con 9 colonne:

| ID | Tipo | Descrizione | Attività/Processo | Probabilità | Gravità | Livello di Rischio | Azioni di Trattamento | Responsabile | Scadenza |
|----|------|-------------|-------------------|-------------|---------|-------------------|----------------------|--------------|----------|

#### Corpo della Tabella (con sintassi corretta docxtpl)

**STRUTTURA ESATTA DA SEGUIRE NEL TEMPLATE WORD:**

```
Riga 1 (Intestazione tabella):
| ID | Tipo | Descrizione | Attività/Processo | Prob. | Grav. | Livello Rischio | Azioni di Trattamento | Responsabile | Scadenza |

Riga 2 (Inizio ciclo FOR - usa {%tr for}):
{%tr for rischio in rischi %}

Riga 3 (Dati - questa riga viene ripetuta per ogni rischio):
| {{ rischio.id }} | {{ rischio.tipo }} | {{ rischio.descrizione }} | {{ rischio.attivita_processo }} | {{ rischio.probabilita }} | {{ rischio.gravita }} | {% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }} | {{ rischio.azioni_trattamento }} | {{ rischio.responsabile }} | {{ rischio.scadenza }} |

Riga 4 (Fine ciclo FOR - usa {%tr endfor}):
{%tr endfor %}
```

**⚠️ ATTENZIONE - ERRORI COMUNI DA EVITARE:**

❌ **SBAGLIATO** (non funziona):
```
{% for rischio in rischi %}
| {{ rischio.id }} | ... |
{% endfor %}
```

✅ **CORRETTO** (funziona):
```
{%tr for rischio in rischi %}
| {{ rischio.id }} | ... |
{%tr endfor %}
```

**Note importanti:**
- `{%tr for %}` e `{%tr endfor %}` devono essere su **righe separate** nella tabella
- Il tag `{% cellbg rischio.codice_colore %}` va **ALL'INIZIO della cella** che vuoi colorare
- Il codice colore deve essere in formato esadecimale SENZA # (es. "FF0000" per rosso)

---

### 9. **Sezione 7 - Flusso Operativo**

Crea una tabella per le fasi del processo:

```
┌───────┬─────────────────┬─────────────────────────────────────────────┐
│ Fase  │ Nome            │ Descrizione                                 │
├───────┼─────────────────┼─────────────────────────────────────────────┤
{%tr for fase in fasi_flusso %}
│ {{ fase.fase }} │ {{ fase.nome }} │ {{ fase.descrizione }} │
{%tr endfor %}
└───────┴─────────────────┴─────────────────────────────────────────────┘
```

---

### 10. **Sezione 8 - Frequenze di Monitoraggio**

```
8. FREQUENZE DI MONITORAGGIO

Le frequenze di monitoraggio sono definite in base al livello di rischio:

• Rischi Estremi: {{ frequenze_monitoraggio.rischi_estremi }}
• Rischi Alti: {{ frequenze_monitoraggio.rischi_alti }}
• Rischi Medi: {{ frequenze_monitoraggio.rischi_medi }}
• Rischi Bassi: {{ frequenze_monitoraggio.rischi_bassi }}
• Opportunità: {{ frequenze_monitoraggio.opportunita }}
```

---

### 11. **Sezione 9 - Record Correlati**

```
9. RECORD CORRELATI

{{ record_correlati_text }}
```

---

### 12. **Piè di Pagina**

Nel piè di pagina del documento Word, inserisci:

```
{{ nome_azienda }} - Procedura PGQ-06-01 - Rev. {{ revisione }} del {{ data_emissione }}
Pagina {{ pagina }}
```

---

## 🎨 Esempio Visivo della Tabella Rischi nel Template

Ecco come deve apparire la struttura nel file Word (visualizza i tag):

```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ IDENTIFICATIVO | TIPO | DESCRIZIONE | PROCESSO | PROBABILITÀ | GRAVITÀ | LIVELLO | AZIONI | RESPONSABILE | SCADENZA │
├─────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ {%tr for rischio in rischi %}                                                                                   │
├─────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ {{ rischio.id }} | {{ rischio.tipo }} | {{ rischio.descrizione }} | {{ rischio.attivita_processo }} |          │
│ {{ rischio.probabilita }} | {{ rischio.gravita }} | {% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }} │
│ {{ rischio.azioni_trattamento }} | {{ rischio.responsabile }} | {{ rischio.scadenza }}                           │
├─────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ {%tr endfor %}                                                                                                  │
└─────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 🔧 Come Inserire i Tag in Word

### Metodo 1: Copia e Incolla Diretto

1. Apri Microsoft Word
2. Crea una nuova tabella
3. Copia i tag Jinja2 esattamente come mostrati sopra
4. Incellali nelle celle appropriate

### Metodo 2: Usare un Template Preesistente

1. Scarica un template Word vuoto
2. Aggiungi le tabelle con la struttura indicata
3. Inserisci i tag Jinja2 nelle celle

---

## ✅ Verifica del Template

Prima di utilizzare lo script, verifica che:

1. ✅ Tutti i tag `{%tr for %}` e `{%tr endfor %}` siano su righe separate della tabella
2. ✅ I tag `{{ variabile }}` siano scritti correttamente (due parentesi graffe)
3. ✅ I codici colore non abbiano il simbolo `#` (usa `FF0000`, non `#FF0000`)
4. ✅ Il file sia salvato come `.docx` (non `.doc` o altri formati)
5. ✅ Non ci siano spazi extra nei tag (usa `{{ rischio.id }}`, non `{{ rischio.id  }}`)

---

## 🧪 Test del Template

Per testare il template:

1. Salva il file come `procedura_template.docx` nella stessa cartella dello script
2. Esegui lo script:
   ```bash
   python generatore_procedura_rischi_v2.py
   ```
3. Verifica il documento generato:
   - Controlla che la tabella dei rischi abbia tutte le righe
   - Verifica che le celle "Livello di Rischio" siano colorate correttamente
   - Controlla che tutti i dati aziendali siano stati inseriti

---

## 🎯 Colori Supportati per le Celle

Ecco i codici colore utilizzati dallo script:

| Livello | Codice Colore | Colore Visivo |
|---------|---------------|---------------|
| Basso | `00FF00` | 🟢 Verde |
| Medio | `FFFF00` | 🟡 Giallo |
| Alto | `FFA500` | 🟠 Arancione |
| Estremo | `FF0000` | 🔴 Rosso |

---

## 📞 Supporto

In caso di problemi:

1. Controlla il file di log `generatore_procedura.log`
2. Verifica che tutti i tag siano scritti correttamente
3. Assicurati che la struttura della tabella segua esattamente le istruzioni
4. Prova con un template minimale per isolare il problema

---

## 📄 Esempio di Template Minimale

Per iniziare, puoi creare un template minimale con solo la tabella dei rischi:

```
PROCEDURA GESTIONE RISCHI

Azienda: {{ nome_azienda }}
Data: {{ data }}

TABELLA RISCHI:

[Tabella con intestazione]
{%tr for rischio in rischi %}
{{ rischio.id }} | {{ rischio.descrizione }} | {% cellbg rischio.codice_colore %}{{ rischio.livello_rischio }}
{%tr endfor %}
```

Una volta verificato che funziona, espandi il template con tutte le sezioni.

---

**Buon lavoro!** 🚀
