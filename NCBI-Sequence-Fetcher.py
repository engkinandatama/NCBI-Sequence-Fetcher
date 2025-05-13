import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
import time
import threading
import pandas as pd

# Import untuk parsing query string
from urllib.parse import parse_qs  # ⬅️ Perbaikan di sini

class NCBISequenceFetcher:
    def __init__(self, root):
        self.root = root
        self.root.title("NCBI Sequence Fetcher")
        self.root.geometry("900x700")

        # Variables
        self.output_folder = tk.StringVar(value=os.path.expanduser("~"))
        self.ncbi_url = tk.StringVar()
        self.report_type = tk.StringVar(value="fasta")
        self.batch_mode = tk.BooleanVar(value=False)
        self.running = False
        self.filename_template = tk.StringVar(value="{accession}_{organism}.{ext}")

        # Batch state management
        self.completed_urls = []
        self.batch_state_file = os.path.join(self.output_folder.get(), "batch_state.json")
        self.metadata_file = os.path.join(self.output_folder.get(), "ncbi_metadata.xlsx")

        # Metadata columns
        self.metadata_columns = [
            'Accession', 'Version', 'Strain', 'Organism', 'Taxonomy',
            'Country', 'Collection_Date', 'Collected_By', 'Isolation_Source',
            'Product', 'Definition', 'Length', 'Filename', 'Downloaded'
        ]
        self.init_metadata()

        # UI Setup
        self.setup_ui()
        
    def init_metadata(self):
        """Initialize or load metadata file"""
        if os.path.exists(self.metadata_file):
            try:
                self.metadata_df = pd.read_excel(self.metadata_file)
            except Exception:
                self.metadata_df = pd.DataFrame(columns=self.metadata_columns)
        else:
            self.metadata_df = pd.DataFrame(columns=self.metadata_columns)

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        ttk.Label(header_frame, text="NCBI Sequence Fetcher", font=('Helvetica', 14, 'bold')).pack(side=tk.LEFT)
        ttk.Checkbutton(header_frame, text="Batch Mode", variable=self.batch_mode,
                         command=self.toggle_batch_mode).pack(side=tk.RIGHT, padx=10)

        url_frame = ttk.Frame(main_frame)
        url_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(url_frame, text="NCBI URLs:").pack(anchor=tk.W)
        self.url_container = ttk.Frame(url_frame)
        self.url_container.pack(fill=tk.BOTH, expand=True)

        self.url_entry = ttk.Entry(self.url_container, textvariable=self.ncbi_url)
        self.url_entry.pack(fill=tk.X, expand=True)
        self.url_entry.bind('<FocusIn>', lambda e: self.url_entry.select_range(0, tk.END))

        self.batch_url_text = scrolledtext.ScrolledText(self.url_container, height=10, wrap=tk.WORD)
        self.batch_url_text.pack(fill=tk.BOTH, expand=True)
        self.batch_url_text.pack_forget()  # Hidden by default

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(control_frame, text="Format:").pack(side=tk.LEFT)
        ttk.OptionMenu(control_frame, self.report_type, "fasta", "fasta", "genbank").pack(side=tk.LEFT, padx=5)

        ttk.Label(control_frame, text="Filename Template:").pack(side=tk.LEFT, padx=(20, 0))
        ttk.Entry(control_frame, textvariable=self.filename_template, width=30).pack(side=tk.LEFT)

        ttk.Label(control_frame, text="Save To:").pack(side=tk.LEFT, padx=(20, 0))
        ttk.Entry(control_frame, textvariable=self.output_folder, width=30).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(control_frame, text="Browse", command=self.browse_folder).pack(side=tk.LEFT)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        self.download_btn = ttk.Button(button_frame, text="Download", command=self.start_download_threaded)
        self.download_btn.pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Clear", command=self.clear_urls).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Import URLs", command=self.import_urls).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Export URLs", command=self.export_urls).pack(side=tk.LEFT)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal",
                                            mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

        self.progress_label = ttk.Label(main_frame, text="Ready")
        self.progress_label.pack(fill=tk.X, padx=5)

        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, padx=5, pady=5)

    def start_download_threaded(self):
        if self.running:
            return
        threading.Thread(target=self.start_download, daemon=True).start()

    def toggle_batch_mode(self):
        if self.batch_mode.get():
            current_url = self.ncbi_url.get()
            self.url_entry.pack_forget()
            self.batch_url_text.pack(fill=tk.BOTH, expand=True)
            if current_url:
                self.batch_url_text.insert(tk.END, current_url + "\n")
            self.download_btn.config(text="Download All")
        else:
            self.batch_url_text.pack_forget()
            self.url_entry.pack(fill=tk.X, expand=True)
            self.download_btn.config(text="Download")

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)
            self.metadata_file = os.path.join(folder, "ncbi_metadata.xlsx")
            self.batch_state_file = os.path.join(folder, "batch_state.json")
            self.log("Output folder set to: " + folder)

    def clear_urls(self):
        if self.batch_mode.get():
            self.batch_url_text.delete(1.0, tk.END)
        else:
            self.ncbi_url.set("")

    def import_urls(self):
        filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filepath:
            with open(filepath, 'r') as f:
                if self.batch_mode.get():
                    self.batch_url_text.insert(tk.END, f.read())
                else:
                    first_line = f.readline().strip()
                    self.ncbi_url.set(first_line)

    def export_urls(self):
        if not self.batch_mode.get():
            messagebox.showinfo("Info", "Available only in Batch Mode")
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".txt",
                                              filetypes=[("Text files", "*.txt")])
        if filepath:
            with open(filepath, 'w') as f:
                f.write(self.batch_url_text.get(1.0, tk.END))
            self.log(f"URLs exported to: {filepath}")

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def update_progress(self, current, total):
        self.progress_var.set(current)
        self.progress_bar["maximum"] = total
        self.progress_label.config(text=f"Processed {current}/{total}")
        self.status_var.set(f"Processing {current}/{total}")
        self.root.update_idletasks()

    def update_status(self, status):
        self.root.after(0, self.status_var.set, status)

    def start_download(self):
        if self.running:
            return
        try:
            self.running = True
            if self.batch_mode.get():
                self.download_batch()
            else:
                self.download_single()
        except Exception as e:
            self.root.after(0, messagebox.showerror, "Error", f"Operation failed: {str(e)}")
            self.log(f"ERROR: {str(e)}")
        finally:
            self.running = False
            self.root.after(0, self.update_status, "Ready")

    def download_single(self):
        url = self.ncbi_url.get().strip()
        if not url:
            self.root.after(0, messagebox.showwarning, "Warning", "Please enter a valid NCBI URL")
            return
        try:
            start_time = time.time()
            self.log(f"STARTING: {url}")
            filename = self.process_url(url)
            duration = time.time() - start_time
            self.log(f"COMPLETED in {duration:.2f}s: {filename}")
            self.root.after(0, messagebox.showinfo, "Success", "Download completed!")
        except Exception as e:
            self.log(f"FAILED: {str(e)}")
            raise

    def download_batch(self):
        all_urls = self.get_urls_from_batch()
        if not all_urls:
            self.root.after(0, messagebox.showwarning, "Warning", "No valid URLs found!")
            return

        self.completed_urls = []
        total = len(all_urls)
        start_time = time.time()
        self.log(f"\n=== BATCH STARTED: {total} URLs ===")
        self.update_progress(0, total)

        try:
            for i, url in enumerate(all_urls, 1):
                if url in self.completed_urls:
                    continue
                try:
                    filename = self.process_url_with_retry(url)
                    self.completed_urls.append(url)
                    self.log(f"COMPLETED: {url} -> {filename}")
                except Exception as e:
                    self.log(f"FAILED: {url} - {str(e)}")
                self.update_progress(i, total)
                self.save_batch_state(self.completed_urls, all_urls)
                time.sleep(1)  # Rate limiting
        except Exception as e:
            self.log(f"BATCH ERROR: {str(e)}")
            raise
        finally:
            duration = time.time() - start_time
            success = len(self.completed_urls)
            self.log(f"=== BATCH COMPLETED: {success}/{total} in {duration:.2f}s ===")
            if success == total:
                self.clear_batch_state()
                self.root.after(0, messagebox.showinfo, "Complete", f"Successfully processed {total} URLs")
            else:
                self.root.after(0, messagebox.showinfo, "Partial Complete",
                                f"Processed {success}/{total} URLs. Failed {total - success}.")

    def process_url_with_retry(self, url, max_retries=3):
        for attempt in range(max_retries):
            try:
                return self.process_url(url)
            except Exception as e:
                if attempt == max_retries - 1:
                    raise
                wait_time = 2 ** attempt
                self.log(f"Retry {attempt+1} for {url} in {wait_time}s...")
                time.sleep(wait_time)

    def get_urls_from_batch(self):
        text = self.batch_url_text.get(1.0, tk.END).strip()
        urls = [url.strip() for url in text.split('\n') if url.strip()]
        valid_urls = []
        for url in urls:
            try:
                if '/nuccore/' in url:
                    accession_id = url.split('/nuccore/')[-1].split('?')[0].strip()
                    base_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id= {accession_id}&rettype={self.report_type.get()}&retmode=text"
                    valid_urls.append(base_url)
                elif url.startswith(("https://www.ncbi.nlm.nih.gov/ ", "http://www.ncbi.nlm.nih.gov/")):
                    valid_urls.append(url)
                else:
                    self.log(f"Invalid URL skipped: {url}")
            except Exception as e:
                self.log(f"Error parsing URL: {url} - {str(e)}")
        return valid_urls

    def process_url(self, url):
        """Process single URL into FASTA or GenBank format"""
        # Parse accession ID from any kind of NCBI URL
        if '/nuccore/' in url:
            accession_id = url.split('/nuccore/')[-1].split('?')[0].strip()
        elif 'id=' in url:
            query_string = url.split('?', 1)[1]
            params = parse_qs(query_string)
            accession_id = params.get('id', [''])[0].strip()
        elif url.startswith(("https://www.ncbi.nlm.nih.gov/ ", "http://www.ncbi.nlm.nih.gov/")):
            parts = url.split('/')
            accession_id = parts[-1].split("?")[0].strip()
        else:
            raise ValueError("Invalid NCBI URL format")

        if not accession_id:
            raise ValueError("Could not extract accession ID")

        ext = self.report_type.get()
        api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id= {accession_id}&rettype={ext}&retmode=text"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
            'Accept': 'text/plain'
        }

        response = requests.get(api_url, headers=headers, timeout=15)
        response.raise_for_status()

        # Validate response format
        if ext == "fasta" and not response.text.startswith('>'):
            raise ValueError("Invalid FASTA format")
        elif ext == "gb" and not response.text.startswith('LOCUS'):
            raise ValueError("Invalid GenBank format")

        metadata = self.extract_metadata(accession_id)
        filename = self.save_data(response.text, metadata, ext)
        self.save_metadata(metadata, filename)
        return filename

    def extract_metadata(self, accession_id):
        metadata = {
            'Accession': accession_id,
            'Version': 'NA',
            'Strain': 'NA',
            'Organism': 'NA',
            'Taxonomy': 'NA',
            'Country': 'NA',
            'Collection_Date': 'NA',
            'Collected_By': 'NA',
            'Isolation_Source': 'NA',
            'Product': 'NA',
            'Definition': 'NA',
            'Length': 'NA',
            'Downloaded': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        try:
            api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id= {accession_id}&rettype=gb&retmode=text"
            response = requests.get(api_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
            response.raise_for_status()
            gb_data = response.text
            lines = gb_data.split('\n')

            for line in lines:
                if line.startswith('VERSION'):
                    metadata['Version'] = line.split()[1].strip()
                elif line.startswith('DEFINITION'):
                    metadata['Definition'] = line[12:].strip()
                elif line.startswith('  ORGANISM'):
                    metadata['Organism'] = line[12:].strip()
                elif '/strain=' in line:
                    metadata['Strain'] = self.extract_value(line, 'strain')
                elif '/country=' in line or '/geo_loc_name=' in line:
                    field = 'country' if '/country=' in line else 'geo_loc_name'
                    metadata['Country'] = self.extract_value(line, field)
                elif '/collection_date=' in line:
                    metadata['Collection_Date'] = self.extract_value(line, 'collection_date')
                elif '/collected_by=' in line:
                    metadata['Collected_By'] = self.extract_value(line, 'collected_by')
                elif '/isolation_source=' in line:
                    metadata['Isolation_Source'] = self.extract_value(line, 'isolation_source')
                elif '/product=' in line:
                    metadata['Product'] = self.extract_value(line, 'product')
                elif line.startswith('LOCUS'):
                    parts = line.split()
                    if len(parts) >= 3:
                        metadata['Length'] = parts[2] + 'bp'
        except Exception as e:
            self.log(f"Metadata warning: {str(e)}")
        return metadata

    def extract_value(self, line, field_name):
        patterns = [f'/{field_name}=\"', f'/{field_name}=']
        for pattern in patterns:
            start = line.find(pattern)
            if start != -1:
                start += len(pattern)
                end = line.find('"', start) if '"' in pattern else min(
                    line.find(' ', start),
                    line.find('\n', start),
                    line.find(';', start),
                    len(line)
                )
                if end == -1:
                    end = len(line)
                return line[start:end].strip()
        return 'NA'

    def save_data(self, data, metadata, ext):
        try:
            filename = self.filename_template.get().format(
                accession=metadata.get('Accession', 'unknown'),
                organism=metadata.get('Organism', 'unknown').replace(' ', '_'),
                strain=metadata.get('Strain', 'unknown'),
                product=metadata.get('Product', 'unknown'),
                length=metadata.get('Length', 'unknown'),
                date=datetime.now().strftime('%Y%m%d'),
                ext=ext
            ).replace('/', '_').replace('\\', '_')
            filename = "".join(c if c.isalnum() or c in ('_', '-', '.') else '_' for c in filename)
            filepath = os.path.join(self.output_folder.get(), filename)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(data)
            self.log(f"Saved: {filename}")
            return filename
        except Exception as e:
            raise ValueError(f"Filename generation failed: {str(e)}")

    def save_metadata(self, metadata, filename):
        try:
            metadata['Filename'] = filename
            new_row = pd.DataFrame([metadata])
            self.metadata_df = pd.concat([self.metadata_df, new_row], ignore_index=True)
            self.metadata_df.to_excel(self.metadata_file, index=False)
        except Exception as e:
            self.log(f"Metadata error: {str(e)}")

    def load_batch_state(self):
        if not self.batch_mode.get():
            return
        if os.path.exists(self.batch_state_file):
            try:
                with open(self.batch_state_file, 'r') as f:
                    state = json.load(f)
                    if state.get('pending_urls'):
                        self.batch_url_text.insert(tk.END, '\n'.join(state['pending_urls']))
                        self.log("Loaded previous batch progress. Click 'Download All' to resume.")
            except Exception as e:
                self.log(f"Error loading batch state: {str(e)}")

    def save_batch_state(self, completed_urls, all_urls):
        pending = [url for url in all_urls if url not in completed_urls]
        try:
            with open(self.batch_state_file, 'w') as f:
                json.dump({
                    "pending_urls": pending,
                    "output_folder": self.output_folder.get()
                }, f)
        except Exception as e:
            self.log(f"Error saving batch state: {str(e)}")

    def clear_batch_state(self):
        if os.path.exists(self.batch_state_file):
            try:
                os.remove(self.batch_state_file)
            except:
                pass

if __name__ == "__main__":
    root = tk.Tk()
    app = NCBISequenceFetcher(root)
    root.mainloop()