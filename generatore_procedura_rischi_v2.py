#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script per la generazione automatica della Procedura per la Gestione dei Rischi e delle Opportunità
Integrata per le norme ISO 9001:2015, ISO 14001:2015, ISO 45001:2018

Autore: Consulente Sistemi di Gestione Integrati
Versione: 2.0
Data: 2024

Miglioramenti v2.0:
- Sintassi corretta docxtpl per tabelle dinamiche ({%tr for %} ... {%tr endfor %})
- Supporto colorazione sfondo celle con {% cellbg %}
- Supporto InlineImage per logo aziendale
- Opzione export PDF (tramite docx2pdf)
- Migliore gestione eccezioni e logging
"""

from docxtpl import DocxTemplate, InlineImage
from datetime import datetime
import os
import sys
import logging
from typing import Dict, List, Any, Optional
from pathlib import Path

# Configurazione logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('generatore_procedura.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class GeneratoreProceduraRischi:
    """Classe per generare la procedura di gestione rischi e opportunità ISO integrata"""
    
    def __init__(self, template_path: str = "procedura_template.docx"):
        """
        Inizializza il generatore con il percorso del template
        
        Args:
            template_path: Percorso del file template Word con segnaposto Jinja2
        """
        self.template_path = template_path
        self.doc = None
        self.output_path = None
        
    def carica_template(self) -> bool:
        """
        Carica il template Word
        
        Returns:
            bool: True se caricato con successo, False altrimenti
        """
        try:
            if not os.path.exists(self.template_path):
                logger.error(f"Template non trovato: {self.template_path}")
                logger.info("Suggerimento: Creare il template seguendo le istruzioni in 'ISTRUZIONI_TEMPLATE.md'")
                return False
            
            self.doc = DocxTemplate(self.template_path)
            logger.info(f"Template caricato con successo: {self.template_path}")
            return True
            
        except Exception as e:
            logger.error(f"Errore nel caricamento del template: {str(e)}", exc_info=True)
            return False
    
    @staticmethod
    def calcola_livello_rischio(probabilita: int, gravita: int) -> Dict[str, Any]:
        """
        Calcola il livello di rischio basato sulla matrice 5x5
        
        Args:
            probabilita: Valore da 1 a 5
            gravita: Valore da 1 a 5
            
        Returns:
            Dict con livello, colore e descrizione
        """
        punteggio = probabilita * gravita
        
        if punteggio <= 4:
            return {
                "livello": "Basso",
                "colore": "Verde",
                "codice_colore": "00FF00",
                "azione": "Accettabile - Monitorare periodicamente"
            }
        elif punteggio <= 9:
            return {
                "livello": "Medio",
                "colore": "Giallo",
                "codice_colore": "FFFF00",
                "azione": "Richiede attenzione - Implementare controlli"
            }
        elif punteggio <= 16:
            return {
                "livello": "Alto",
                "colore": "Arancione",
                "codice_colore": "FFA500",
                "azione": "Non accettabile - Azioni correttive immediate"
            }
        else:
            return {
                "livello": "Estremo",
                "colore": "Rosso",
                "codice_colore": "FF0000",
                "azione": "Critico - Fermare attività e intervenire immediatamente"
            }
    
    @staticmethod
    def prepara_rischi_per_template(rischi: List[Dict], master: Optional[DocxTemplate] = None) -> List[Dict]:
        """
        Prepara la lista dei rischi aggiungendo il calcolo automatico del livello
        e gestendo le immagini inline se presenti
        
        Args:
            rischi: Lista di dizionari con i dati dei rischi
            master: Oggetto DocxTemplate per gestire InlineImage
            
        Returns:
            Lista di dizionari con dati arricchiti per il template
        """
        rischi_preparati = []
        
        for rischio in rischi:
            rischio_copy = rischio.copy()
            
            # Calcola livello di rischio se non presente
            if "livello_rischio" not in rischio_copy:
                calcolo = GeneratoreProceduraRischi.calcola_livello_rischio(
                    rischio_copy.get("probabilita", 1),
                    rischio_copy.get("gravita", 1)
                )
                rischio_copy["livello_rischio"] = calcolo["livello"]
                rischio_copy["colore_rischio"] = calcolo["colore"]
                rischio_copy["codice_colore"] = calcolo["codice_colore"]
                rischio_copy["azione_richiesta"] = calcolo["azione"]
            
            # Formatta la scadenza
            if "scadenza" in rischio_copy and isinstance(rischio_copy["scadenza"], datetime):
                rischio_copy["scadenza"] = rischio_copy["scadenza"].strftime("%d/%m/%Y")
            
            # Gestisci immagine inline se presente nel rischio (es. allegati)
            if "immagine_path" in rischio_copy and master is not None:
                try:
                    rischio_copy["immagine"] = InlineImage(
                        master, 
                        rischio_copy["immagine_path"],
                        width=30000000  # 30mm in EMU
                    )
                except Exception as e:
                    logger.warning(f"Impossibile caricare immagine per rischio {rischio_copy.get('id', 'N/A')}: {str(e)}")
                    rischio_copy["immagine"] = None
            
            rischi_preparati.append(rischio_copy)
        
        return rischi_preparati
    
    def prepara_context(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """
        Prepara il context completo per il rendering, inclusi logo e immagini
        
        Args:
            context: Dizionario con i dati originali
            
        Returns:
            Dizionario pronto per il rendering
        """
        context_preparato = context.copy()
        
        # Prepara i rischi se presenti
        if "rischi" in context_preparato:
            context_preparato["rischi"] = self.prepara_rischi_per_template(
                context_preparato["rischi"], 
                self.doc
            )
        
        # Aggiungi data corrente se non presente
        if "data" not in context_preparato:
            context_preparato["data"] = datetime.now().strftime("%d/%m/%Y")
        
        if "data_anno" not in context_preparato:
            context_preparato["data_anno"] = datetime.now().strftime("%Y")
        
        # Gestisci logo aziendale se presente
        if "logo_path" in context_preparato and os.path.exists(context_preparato["logo_path"]):
            try:
                context_preparato["logo"] = InlineImage(
                    self.doc,
                    context_preparato["logo_path"],
                    width=60000000,  # 60mm in EMU (circa 3cm)
                    height=30000000  # 30mm in EMU (circa 1.5cm)
                )
                logger.info(f"Logo aziendale caricato: {context_preparato['logo_path']}")
            except Exception as e:
                logger.warning(f"Impossibile caricare logo aziendale: {str(e)}")
                context_preparato["logo"] = None
        else:
            context_preparato["logo"] = None
        
        # Formatta riferimenti normativi come stringa se è una lista
        if "riferimenti_normativi" in context_preparato and isinstance(context_preparato["riferimenti_normativi"], list):
            context_preparato["riferimenti_normativi_text"] = "\n".join(
                f"• {ref}" for ref in context_preparato["riferimenti_normativi"]
            )
        
        # Formatta record correlati come stringa se è una lista
        if "record_correlati" in context_preparato and isinstance(context_preparato["record_correlati"], list):
            context_preparato["record_correlati_text"] = "\n".join(
                f"• {rec}" for rec in context_preparato["record_correlati"]
            )
        
        return context_preparato
    
    def genera_documento(self, context: Dict[str, Any], output_path: str = None, converti_pdf: bool = False) -> bool:
        """
        Genera il documento Word finale con opzione di conversione PDF
        
        Args:
            context: Dizionario con tutti i dati per il template
            output_path: Percorso del file output (opzionale)
            converti_pdf: Se True, tenta di convertire in PDF dopo la generazione
            
        Returns:
            bool: True se generato con successo, False altrimenti
        """
        try:
            if self.doc is None:
                if not self.carica_template():
                    return False
            
            # Prepara il context completo
            context_preparato = self.prepara_context(context)
            
            # Genera il nome del file output se non specificato
            if output_path is None:
                nome_azienda = context.get("nome_azienda", "Azienda").replace(" ", "_").replace(".", "")
                output_path = f"Procedura_Gestione_Rischi_{nome_azienda}_{datetime.now().strftime('%Y%m%d')}.docx"
            
            self.output_path = output_path
            
            # Renderizza il template
            logger.info("Rendering del template in corso...")
            self.doc.render(context_preparato)
            
            # Salva il documento
            logger.info(f"Salvataggio documento: {output_path}")
            self.doc.save(output_path)
            
            logger.info(f"Documento generato con successo: {output_path}")
            
            # Conversione opzionale in PDF
            if converti_pdf:
                self._converti_in_pdf(output_path)
            
            return True
            
        except Exception as e:
            logger.error(f"Errore nella generazione del documento: {str(e)}", exc_info=True)
            return False
    
    def _converti_in_pdf(self, docx_path: str) -> bool:
        """
        Tenta di convertire il documento Word in PDF
        
        Args:
            docx_path: Percorso del file DOCX
            
        Returns:
            bool: True se conversione riuscita, False altrimenti
        """
        try:
            from docx2pdf import convert
            
            pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
            logger.info(f"Conversione in PDF in corso: {pdf_path}")
            
            convert(docx_path, pdf_path)
            
            logger.info(f"PDF generato con successo: {pdf_path}")
            return True
            
        except ImportError:
            logger.warning("Modulo 'docx2pdf' non installato. Installare con: pip install docx2pdf")
            logger.info("Il documento Word è stato generato correttamente, ma la conversione PDF non è stata eseguita.")
            return False
        except Exception as e:
            logger.error(f"Errore nella conversione PDF: {str(e)}")
            logger.info("Nota: La conversione PDF richiede Microsoft Word o LibreOffice installati sul sistema.")
            return False
    
    def valida_context(self, context: Dict[str, Any]) -> List[str]:
        """
        Valida che il context contenga tutti i campi obbligatori
        
        Args:
            context: Dizionario da validare
            
        Returns:
            Lista di messaggi di errore (vuota se valido)
        """
        errori = []
        
        campi_obbligatori = [
            "nome_azienda",
            "responsabile",
            "revisione",
            "data_emissione"
        ]
        
        for campo in campi_obbligatori:
            if campo not in context or not context[campo]:
                errori.append(f"Campo obbligatorio mancante: {campo}")
        
        # Valida la struttura dei rischi se presenti
        if "rischi" in context:
            for i, rischio in enumerate(context["rischi"]):
                if "descrizione" not in rischio:
                    errori.append(f"Rischio {i+1}: manca la descrizione")
                if "probabilita" not in rischio or not 1 <= rischio.get("probabilita", 0) <= 5:
                    errori.append(f"Rischio {i+1}: probabilità non valida (deve essere 1-5)")
                if "gravita" not in rischio or not 1 <= rischio.get("gravita", 0) <= 5:
                    errori.append(f"Rischio {i+1}: gravità non valida (deve essere 1-5)")
                
                # Validazione opzionale per immagini
                if "immagine_path" in rischio:
                    if not os.path.exists(rischio["immagine_path"]):
                        errori.append(f"Rischio {i+1}: file immagine non trovato: {rischio['immagine_path']}")
        
        # Validazione logo se specificato
        if "logo_path" in context:
            if not os.path.exists(context["logo_path"]):
                errori.append(f"File logo non trovato: {context['logo_path']}")
        
        return errori
    
    def get_info_documento(self) -> Dict[str, Any]:
        """
        Restituisce informazioni sul documento generato
        
        Returns:
            Dict con informazioni sul documento
        """
        return {
            "template_path": self.template_path,
            "output_path": self.output_path,
            "template_caricato": self.doc is not None
        }


def crea_context_esempio() -> Dict[str, Any]:
    """
    Crea un esempio di context per testare lo script
    
    Returns:
        Dict con dati di esempio completi
    """
    return {
        # Dati Aziendali
        "nome_azienda": "Industria Italiana S.p.A.",
        "indirizzo": "Via Roma 123, 20100 Milano (MI)",
        "partita_iva": "IT12345678901",
        "telefono": "+39 02 1234567",
        "email": "info@industriaitaliana.it",
        "sito_web": "www.industriaitaliana.it",
        # "logo_path": "logo_azienda.png",  # Decommentare se si ha un logo
        
        # Dati Documento
        "codice_documento": "PGQ-06-01",
        "revisione": "03",
        "data_emissione": "15/01/2024",
        "data_anno": "2024",
        "pagina": "1 di 8",
        
        # Responsabili
        "responsabile": "Ing. Marco Rossi",
        "ruolo_responsabile": "Responsabile Sistema di Gestione Integrato",
        "approvato_da": "Dott. Giuseppe Bianchi",
        "ruolo_approvatore": "Direttore Generale",
        "data_approvazione": "20/01/2024",
        
        # Riferimenti Normativi
        "riferimenti_normativi": [
            "UNI EN ISO 9001:2015 - Paragrafo 6.1",
            "UNI EN ISO 14001:2015 - Paragrafo 6.1",
            "UNI EN ISO 45001:2018 - Paragrafo 6.1",
            "D.Lgs. 81/2008 - Valutazione dei Rischi"
        ],
        
        # Definizioni
        "definizioni": {
            "rischio": "Effetto dell'incertezza sugli obiettivi",
            "opportunita": "Circostanza favorevole al raggiungimento degli obiettivi",
            "parte_interessata": "Persona o organizzazione che può influenzare o essere influenzata",
            "valutazione_rischio": "Processo complessivo di identificazione, analisi e valutazione del rischio"
        },
        
        # Matrice di Valutazione
        "scala_probabilita": [
            {"valore": 1, "descrizione": "Molto improbabile", "frequenza": "Una volta ogni 10 anni"},
            {"valore": 2, "descrizione": "Improbabile", "frequenza": "Una volta ogni 5 anni"},
            {"valore": 3, "descrizione": "Possibile", "frequenza": "Una volta all'anno"},
            {"valore": 4, "descrizione": "Probabile", "frequenza": "Una volta al mese"},
            {"valore": 5, "descrizione": "Molto probabile", "frequenza": "Una volta a settimana o più"}
        ],
        
        "scala_gravita": [
            {"valore": 1, "descrizione": "Insignificante", "impatto": "Nessun danno o impatto minimo"},
            {"valore": 2, "descrizione": "Minore", "impatto": "Danno lieve, trattamento semplice"},
            {"valore": 3, "descrizione": "Moderato", "impatto": "Danno con trattamento medico, impatto ambientale limitato"},
            {"valore": 4, "descrizione": "Maggiore", "impatto": "Infortunio grave, impatto ambientale significativo"},
            {"valore": 5, "descrizione": "Catastrofico", "impatto": "Morte, impatto ambientale grave e permanente"}
        ],
        
        # Elenco Rischi e Opportunità
        "rischi": [
            {
                "id": "R001",
                "descrizione": "Malfunzionamento macchinari di produzione",
                "attivita_processo": "Produzione",
                "tipo": "Rischio",
                "probabilita": 3,
                "gravita": 4,
                "azioni_trattamento": "Manutenzione preventiva programmata, formazione operatori",
                "responsabile": "Resp. Manutenzione",
                "scadenza": "31/12/2024"
                # "immagine_path": "foto_macchinario.jpg",  # Opzionale
            },
            {
                "id": "R002",
                "descrizione": "Esposizione a sostanze chimiche pericolose",
                "attivita_processo": "Stoccaggio materiali",
                "tipo": "Rischio",
                "probabilita": 2,
                "gravita": 5,
                "azioni_trattamento": "DPI obbligatori, ventilazione adeguata, formazione specifica",
                "responsabile": "Resp. HSE",
                "scadenza": "30/06/2024"
            },
            {
                "id": "R003",
                "descrizione": "Ritardo fornitori materiali critici",
                "attivita_processo": "Approvvigionamento",
                "tipo": "Rischio",
                "probabilita": 4,
                "gravita": 3,
                "azioni_trattamento": "Qualifica fornitori alternativi, scorte di sicurezza",
                "responsabile": "Resp. Acquisti",
                "scadenza": "28/02/2024"
            },
            {
                "id": "O001",
                "descrizione": "Nuova tecnologia per riduzione consumi energetici",
                "attivita_processo": "Gestione energetica",
                "tipo": "Opportunità",
                "probabilita": 4,
                "gravita": 3,
                "azioni_trattamento": "Valutazione investimento, progetto pilota",
                "responsabile": "Resp. Energia",
                "scadenza": "30/09/2024"
            },
            {
                "id": "O002",
                "descrizione": "Espansione mercato estero",
                "attivita_processo": "Commerciale",
                "tipo": "Opportunità",
                "probabilita": 3,
                "gravita": 4,
                "azioni_trattamento": "Studio di mercato, partnership locali",
                "responsabile": "Direttore Commerciale",
                "scadenza": "31/12/2024"
            }
        ],
        
        # Flusso Operativo
        "fasi_flusso": [
            {"fase": 1, "nome": "Identificazione", "descrizione": "Rilevamento rischi e opportunità da tutte le fonti"},
            {"fase": 2, "nome": "Valutazione", "descrizione": "Analisi probabilità e gravità, calcolo livello rischio"},
            {"fase": 3, "nome": "Trattamento", "descrizione": "Definizione azioni per mitigare rischi o sfruttare opportunità"},
            {"fase": 4, "nome": "Monitoraggio", "descrizione": "Verifica efficacia azioni e aggiornamento periodico"},
            {"fase": 5, "nome": "Riesame", "descrizione": "Revisione annuale da parte della Direzione"}
        ],
        
        # Frequenze di Monitoraggio
        "frequenze_monitoraggio": {
            "rischi_estremi": "Mensile",
            "rischi_alti": "Trimestrale",
            "rischi_medi": "Semestrale",
            "rischi_bassi": "Annuale",
            "opportunita": "Trimestrale"
        },
        
        # Record Correlati
        "record_correlati": [
            "Registro Rischi e Opportunità (FRQ-06-01)",
            "Verbale Riesame della Direzione (FRQ-09-03)",
            "Piano di Azione Correttiva (FRQ-10-01)",
            "Registro Non Conformità (FRQ-10-02)"
        ]
    }


def main():
    """Funzione principale per eseguire lo script"""
    
    print("=" * 80)
    print("GENERATORE PROCEDURA GESTIONE RISCHI E OPPORTUNITÀ ISO 9001/14001/45001")
    print("Versione 2.0 - Con supporto tabelle dinamiche, colorazione e export PDF")
    print("=" * 80)
    print()
    
    # Crea istanza del generatore
    generatore = GeneratoreProceduraRischi("procedura_template.docx")
    
    # Carica il context di esempio (o personalizzato)
    context = crea_context_esempio()
    
    # Valida il context
    logger.info("Validazione dati in corso...")
    errori = generatore.valida_context(context)
    
    if errori:
        print("\n⚠️  ERRORI DI VALIDAZIONE:")
        for errore in errori:
            print(f"  - {errore}")
        print("\nCorreggere gli errori prima di generare il documento.")
        print("\nSuggerimento: Leggere le istruzioni in 'ISTRUZIONI_TEMPLATE.md' per creare il template corretto.")
        return False
    
    logger.info("✓ Validazione completata con successo")
    print()
    
    # Chiedi all'utente se vuole convertire in PDF
    converti_pdf = input("Vuoi convertire il documento in PDF? (s/n): ").strip().lower() == 's'
    
    if converti_pdf:
        logger.info("Conversione PDF richiesta")
        print("Nota: La conversione PDF richiede Microsoft Word o LibreOffice installati.")
    
    # Genera il documento
    print("\nGenerazione documento in corso...")
    successo = generatore.genera_documento(context, converti_pdf=converti_pdf)
    
    if successo:
        info = generatore.get_info_documento()
        print("\n" + "=" * 80)
        print("DOCUMENTO GENERATO CON SUCCESSO!")
        print("=" * 80)
        print(f"\n📄 File Word creato: {info['output_path']}")
        
        if converti_pdf and os.path.exists(os.path.splitext(info['output_path'])[0] + ".pdf"):
            print(f"📑 File PDF creato: {os.path.splitext(info['output_path'])[0]}.pdf")
        
        print(f"🕐 Data generazione: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"📊 Numero rischi/opportunità inclusi: {len(context['rischi'])}")
        print("\n✓ Procedura pronta per la revisione e l'approvazione")
        print("\n💡 Suggerimento: Verificare che la tabella dei rischi mostri correttamente:")
        print("   - Le righe dinamiche per ogni rischio/opportunità")
        print("   - La colorazione dello sfondo nella colonna 'Livello di Rischio'")
        print("   - Il logo aziendale nell'intestazione (se fornito)")
    else:
        print("\n✗ Errore nella generazione del documento")
        print("\nControllare il file di log 'generatore_procedura.log' per dettagli.")
        return False
    
    return True


if __name__ == "__main__":
    main()
