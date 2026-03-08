import os
import sys
import re
import json
import glob
import datetime
import io
import time
import threading
import queue
import fitz
import easyocr
import requests
from fpdf import FPDF
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

OUTPUT_PDF = 'output_ddt.pdf'
DPI = 300
OLLAMA_URL = 'http://localhost:11434/api/generate'
OLLAMA_MODEL = 'mistral'

_ocr_reader = None

def get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        _ocr_reader = easyocr.Reader(['it', 'en'], gpu=False)
    return _ocr_reader


def pdf_to_text(pdf_path, log):
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text().strip()
        if len(text) > 30:
            pages.append(text)
        else:
            log(f'  Pagina {i+1}: OCR in corso...')
            mat = fitz.Matrix(DPI / 72, DPI / 72)
            pix = page.get_pixmap(matrix=mat)
            results = get_ocr_reader().readtext(pix.tobytes('png'), detail=0, paragraph=True)
            pages.append('\n'.join(results))
    doc.close()
    return '\n'.join(pages)

COLLI_PATTERNS = [
    r'(\d+)\s*(?:colli|pallet|bancali|casse|cartoni|fusti|sacchi|imballi|pedane)',
]

MERCE_PATTERNS = [
    r'(?:descrizione\s+(?:della\s+)?merce)[^\n:]*[:\s]+([^\n]{5,150})',
    r'(?:causale\s+(?:del\s+)?trasporto)[^\n:]*[:\s]+([^\n]{5,150})',
    r'(?:descrizione)[^\n:]*:\s*([^\n]{5,150})',
    r'(?:merce)[^\n:]*:\s*([^\n]{5,150})',
]

COMMESSA_PATTERNS = [
    r'(?:commessa|cod(?:ice)?\s*commessa)[^\n:]*[:\s]+([^\n]{2,80})',
    r'(?:job|progetto)[^\n:]*[:\s]+([^\n]{2,80})',
]

ORDINE_PATTERNS = [
    r'(?:n\.?\s*ordine|num(?:ero)?\s*ordine)[^\n:]*[:\s]+([^\n]{2,80})',
    r'(?:ordine\s+cliente|vs\.?\s*ordine)[^\n:]*[:\s]+([^\n]{2,80})',
]

def extract_with_regex(text):
    result = {'descrizione_merce': None, 'numero_colli': None, 'commessa': None, 'ordine': None}
    for p in COLLI_PATTERNS:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            result['numero_colli'] = m.group(1).strip()
            break
    for p in MERCE_PATTERNS:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            d = m.group(1).strip().rstrip('.,;')
            if len(d) > 3:
                result['descrizione_merce'] = d
                break
    for p in COMMESSA_PATTERNS:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            result['commessa'] = m.group(1).strip().rstrip('.,;')
            break
    for p in ORDINE_PATTERNS:
        m = re.search(p, text, re.IGNORECASE)
        if m:
            result['ordine'] = m.group(1).strip().rstrip('.,;')
            break
    return result


def ollama_available():
    try:
        return requests.get('http://localhost:11434/api/tags', timeout=3).status_code == 200
    except:
        return False


def model_available(name):
    try:
        r = requests.get('http://localhost:11434/api/tags', timeout=3)
        return any(name in m['name'] for m in r.json().get('models', []))
    except:
        return False

def extract_with_ollama(text, log):
    prompt = ('Sei un assistente che analizza DDT italiani.\n' 'Estrai SOLO questi campi e rispondi con JSON puro, senza markdown:\n' '- "descrizione_merce"\n' '- "numero_colli"\n' '- "commessa"\n' '- "ordine"\n' 'Se assente usa null.\n\nTESTO:\n' + text[:2500] + '\n\nJSON:')
    try:
        log('  Ollama Mistral in elaborazione (CPU, attendere)...')
        t0 = time.time()
        resp = requests.post(OLLAMA_URL, json={'model': OLLAMA_MODEL, 'prompt': prompt, 'stream': False}, timeout=300)
        content = re.sub(r'```(?:json)?', '', resp.json().get('response', '')).strip().rstrip('`')
        data = json.loads(content)
        log(f'  Ollama OK ({round(time.time() - t0, 1)}s)')
        return data
    except:
        log('  Ollama: risposta non utilizzabile.')
        return {'descrizione_merce': None, 'numero_colli': None, 'commessa': None, 'ordine': None}


def extract_fields(text, filename, use_ollama, log):
    result = extract_with_regex(text)
    missing = [k for k, v in result.items() if v is None]
    found = [k for k, v in result.items() if v is not None]
    if not missing:
        log('  Regex: tutti i campi trovati.')
        result['metodo'] = 'regex'
    elif use_ollama:
        log(f'  Regex: trovati {found or "nulla"}, mancanti {missing}')
        res = extract_with_ollama(text, log)
        for f in missing:
            if res.get(f):
                result[f] = res[f]
        result['metodo'] = 'regex+ollama' if found else 'ollama'
    else:
        result['metodo'] = 'regex (parziale)'
    result['file'] = os.path.basename(filename)
    return result

def build_pdf_section(records):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Helvetica', 'B', 13)
    pdf.set_fill_color(30, 80, 160)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 11, 'ESTRAZIONE DDT - DOCUMENTI DI TRASPORTO', ln=True, fill=True, align='C')
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(130, 130, 130)
    pdf.cell(0, 6, f"Sessione: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True, align='R')
    pdf.ln(3)
    W = {'file': 40, 'merce': 68, 'colli': 18, 'commessa': 26, 'ordine': 26, 'metodo': 22}
    TOT = sum(W.values())
    pdf.set_font('Helvetica', 'B', 8)
    pdf.set_fill_color(210, 225, 250)
    pdf.set_text_color(0, 0, 0)
    for label, w in [('File PDF', W['file']), ('Descrizione Merce', W['merce']), ('N. Colli', W['colli']), ('Commessa', W['commessa']), ('Ordine', W['ordine']), ('Metodo', W['metodo'])]:
        pdf.cell(w, 9, label, border=1, fill=True, align='C')
    pdf.ln()
    pdf.set_font('Helvetica', '', 7)
    for idx, rec in enumerate(records):
        fill = idx % 2 == 0
        if fill:
            pdf.set_fill_color(245, 248, 255)
        else:
            pdf.set_fill_color(255, 255, 255)
        h = 7
        x0, y0 = pdf.get_x(), pdf.get_y()
        cur_x = x0
        for val, w in [(rec.get('file', '-'), W['file']), (str(rec.get('descrizione_merce') or 'Non trovato'), W['merce']), (str(rec.get('numero_colli') or '-'), W['colli']), (str(rec.get('commessa') or '-'), W['commessa']), (str(rec.get('ordine') or '-'), W['ordine']), (rec.get('metodo', '-'), W['metodo'])]:
            pdf.set_xy(cur_x, y0)
            pdf.multi_cell(w, h, val, border='LR', fill=fill, align='L', max_line_height=h)
            cur_x += w
        pdf.set_xy(x0, y0 + h)
        pdf.cell(TOT, 0, '', border='T')
        pdf.ln()
    pdf.ln(4)
    pdf.set_font('Helvetica', 'B', 9)
    pdf.set_text_color(30, 80, 160)
    pdf.cell(0, 7, f'DDT elaborati: {len(records)}', ln=True)
    return bytes(pdf.output())


def save_output(records, output_path, log):
    new_bytes = build_pdf_section(records)
    if os.path.exists(output_path):
        merged = fitz.open()
        merged.insert_pdf(fitz.open(output_path))
        merged.insert_pdf(fitz.open('pdf', new_bytes))
        merged.save(output_path)
        merged.close()
        log(f'PDF aggiornato: {output_path}')
    else:
        open(output_path, 'wb').write(new_bytes)
        log(f'PDF creato: {output_path}')

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Estrattore DDT')
        self.geometry('740x600')
        self.configure(bg='#f0f4f8')
        self.pdf_files = []
        self.output_path = tk.StringVar(value=os.path.join(os.path.expanduser('~'), 'Desktop', OUTPUT_PDF))
        self.use_ollama = tk.BooleanVar(value=True)
        self.running = False
        self.log_queue = queue.Queue()
        self._build_ui()
        self._check_ollama()
        self.after(100, self._poll_log)

    def _build_ui(self):
        hdr = tk.Frame(self, bg='#1e50a0', pady=10)
        hdr.pack(fill='x')
        tk.Label(hdr, text='Estrattore DDT', font=('Helvetica', 16, 'bold'), bg='#1e50a0', fg='white').pack()
        tk.Label(hdr, text='Merce | Colli | Commessa | Ordine da PDF scansionati', font=('Helvetica', 9), bg='#1e50a0', fg='#c8d8f0').pack()
        body = tk.Frame(self, bg='#f0f4f8', padx=16, pady=12)
        body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=1)
        tk.Label(body, text='PDF da elaborare', font=('Helvetica', 10, 'bold'), bg='#f0f4f8').grid(row=0, column=0, sticky='w', pady=(0, 4))
        ff = tk.Frame(body, bg='#f0f4f8')
        ff.grid(row=1, column=0, sticky='ew')
        self.listbox = tk.Listbox(ff, height=7, selectmode=tk.EXTENDED, font=('Helvetica', 9), relief='solid', bd=1)
        self.listbox.pack(side='left', fill='both', expand=True)
        ttk.Scrollbar(ff, orient='vertical', command=self.listbox.yview).pack(side='right', fill='y')
        bf = tk.Frame(body, bg='#f0f4f8')
        bf.grid(row=2, column=0, sticky='w', pady=6)
        for text, cmd, color in [('+ Aggiungi PDF', self._add_files, '#1e50a0'), ('  Aggiungi cartella', self._add_folder, '#4a7abf'), ('  Rimuovi selezionati', self._remove_sel, '#c0392b')]:
            tk.Button(bf, text=text, command=cmd, bg=color, fg='white', relief='flat', padx=10, pady=4, cursor='hand2').pack(side='left', padx=(0, 6))
        tk.Label(body, text='File di output', font=('Helvetica', 10, 'bold'), bg='#f0f4f8').grid(row=3, column=0, sticky='w', pady=(8, 2))
        of = tk.Frame(body, bg='#f0f4f8')
        of.grid(row=4, column=0, sticky='ew')
        tk.Entry(of, textvariable=self.output_path, font=('Helvetica', 9), relief='solid', bd=1).pack(side='left', fill='x', expand=True)
        tk.Button(of, text='Sfoglia', command=self._browse_out, bg='#5d6d7e', fg='white', relief='flat', padx=8, cursor='hand2').pack(side='left', padx=(6, 0))
        of2 = tk.Frame(body, bg='#f0f4f8')
        of2.grid(row=5, column=0, sticky='w', pady=(10, 4))
        tk.Checkbutton(of2, text='Usa Ollama (Mistral) come fallback AI', variable=self.use_ollama, bg='#f0f4f8', font=('Helvetica', 9)).pack(side='left')
        self.ollama_lbl = tk.Label(of2, text='', font=('Helvetica', 8), bg='#f0f4f8')
        self.ollama_lbl.pack(side='left', padx=8)
        self.run_btn = tk.Button(body, text='▶  Avvia Estrazione', command=self._start, bg='#27ae60', fg='white', font=('Helvetica', 11, 'bold'), relief='flat', padx=16, pady=8, cursor='hand2')
        self.run_btn.grid(row=6, column=0, pady=10, sticky='ew')
        self.progress = ttk.Progressbar(body, mode='determinate')
        self.progress.grid(row=7, column=0, sticky='ew', pady=(0, 6))
        tk.Label(body, text='Log', font=('Helvetica', 9, 'bold'), bg='#f0f4f8').grid(row=8, column=0, sticky='w')
        lf = tk.Frame(body, bg='#f0f4f8')
        lf.grid(row=9, column=0, sticky='nsew')
        body.rowconfigure(9, weight=1)
        self.log_txt = tk.Text(lf, height=7, font=('Courier', 8), state='disabled', relief='solid', bd=1, bg='#1a1a2e', fg='#a8d8a8')
        self.log_txt.pack(side='left', fill='both', expand=True)
        ttk.Scrollbar(lf, orient='vertical', command=self.log_txt.yview).pack(side='right', fill='y')

    def _add_files(self):
        files = filedialog.askopenfilenames(title='Seleziona PDF', filetypes=[('PDF', '*.pdf *.PDF'), ('Tutti', '*.*')])
        for f in files:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))

    def _add_folder(self):
        folder = filedialog.askdirectory(title='Seleziona cartella')
        if folder:
            found = glob.glob(os.path.join(folder, '*.[pP][dD][fF]'))
            added = 0
            for f in found:
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
                    self.listbox.insert(tk.END, os.path.basename(f))
                    added += 1
            self._log(f'Aggiunti {added} PDF da: {folder}')

    def _remove_sel(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            self.pdf_files.pop(idx)

    def _browse_out(self):
        p = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF', '*.pdf')])
        if p:
            self.output_path.set(p)

    def _check_ollama(self):
        def _chk():
            if ollama_available() and model_available(OLLAMA_MODEL):
                self.ollama_lbl.config(text='● Mistral pronto', fg='#27ae60')
            elif ollama_available():
                self.ollama_lbl.config(text='● Ollama attivo, Mistral non trovato', fg='#e67e22')
                self.use_ollama.set(False)
            else:
                self.ollama_lbl.config(text='● Ollama non raggiungibile', fg='#c0392b')
                self.use_ollama.set(False)
        threading.Thread(target=_chk, daemon=True).start()

    def _log(self, msg):
        self.log_queue.put(msg)

    def _poll_log(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_txt.configure(state='normal')
            self.log_txt.insert(tk.END, msg + '\n')
            self.log_txt.see(tk.END)
            self.log_txt.configure(state='disabled')
        self.after(100, self._poll_log)

    def _start(self):
        if self.running:
            return
        if not self.pdf_files:
            messagebox.showwarning('Nessun file', 'Aggiungi almeno un PDF.')
            return
        if not self.output_path.get().strip():
            messagebox.showwarning('Output mancante', 'Specifica il file di output.')
            return
        self.running = True
        self.run_btn.configure(state='disabled', text='Elaborazione in corso...')
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.pdf_files)
        threading.Thread(target=self._run, args=(list(self.pdf_files), self.output_path.get(), self.use_ollama.get()), daemon=True).start()

    def _run(self, pdf_files, output_path, use_ollama):
        self._log(f"\n=== Avvio: {datetime.datetime.now().strftime('%H:%M:%S')} ===")
        records = []
        for i, path in enumerate(pdf_files, 1):
            self._log(f"\n[{i}/{len(pdf_files)}] {os.path.basename(path)}")
            try:
                text = pdf_to_text(path, self._log)
                record = extract_fields(text, path, use_ollama, self._log)
                records.append(record)
                self._log(f"  -> Merce: {record.get('descrizione_merce')} | Colli: {record.get('numero_colli')} | Commessa: {record.get('commessa')} | Ordine: {record.get('ordine')}")
            except Exception as e:
                self._log(f'  ERRORE: {e}')
            self.after(0, lambda v=i: self.progress.configure(value=v))
        if records:
            try:
                save_output(records, output_path, self._log)
                self._log(f'\nFatto! -> {os.path.abspath(output_path)}')
                self.after(0, lambda: messagebox.showinfo('Completato', f"Estrazione completata!\nDDT: {len(records)}\n{os.path.abspath(output_path)}"))
                self.after(0, lambda: os.startfile(os.path.dirname(os.path.abspath(output_path))))
            except Exception as e:
                self._log(f'Errore salvataggio: {e}')
                self.after(0, lambda: messagebox.showerror('Errore', str(e)))
        else:
            self._log('Nessun dato estratto.')
        self.after(0, self._done)

    def _done(self):
        self.running = False
        self.run_btn.configure(state='normal', text='▶  Avvia Estrazione')


if __name__ == '__main__':
    App().mainloop()
