Applicazione desktop per estrarre automaticamente dati da documenti di trasporto (DDT) scansionati in PDF.

Estrae i seguenti campi:

📦 Descrizione merce
🔢 Numero colli
🏷️ Commessa
📋 Ordine
✅ Requisiti
Python 3.8 o superiore → scarica qui
Ollama con Mistral installato → scarica qui
Nessun privilegio amministratore richiesto
🚀 Installazione
Scarica il repository cliccando Code → Download ZIP
Estrai la cartella dove preferisci
Doppio click su installa.bat (solo la prima volta)
▶️ Utilizzo
Doppio click su avvia.bat
Clicca + Aggiungi PDF oppure Aggiungi cartella
Scegli dove salvare il file di output
Clicca ▶ Avvia Estrazione
Al termine si apre automaticamente la cartella con il PDF generato.

🧠 Come funziona
Fase	Dettaglio
Lettura PDF	Estrae testo direttamente se digitale, applica OCR (EasyOCR) se scansionato
Regex	Cerca i campi con pattern specifici per DDT italiani
Ollama Mistral	Usato come fallback AI se la regex non trova tutti i campi
Output	PDF cumulativo — ogni sessione aggiunge i nuovi dati senza cancellare i precedenti
📁 File inclusi
File	Descrizione
estrai_ddt.py	Script principale con GUI
installa.bat	Installa le dipendenze Python (eseguire una volta sola)
avvia.bat	Avvia l'applicazione
📦 Dipendenze Python

Installate automaticamente da installa.bat:

pymupdf  easyocr  fpdf2  requests
