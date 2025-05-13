import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

class NCBISequenceFetcher:
    def __init__(self, root):
        self.root = root
        self.root.title("NCBI Sequence Fetcher")
        self.root.geometry("800x600")
        
        # Variables
        self.output_folder = tk.StringVar(value=os.path.expanduser("~"))
        self.ncbi_url = tk.StringVar()
        self.report_type = tk.StringVar(value="fasta")
        self.batch_mode = tk.BooleanVar(value=False)
        self.running = False
        
        # Metadata setup
        self.metadata_file = os.path.join(self.output_folder.get(), "ncbi_metadata.xlsx")
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
            except:
                self.metadata_df = pd.DataFrame(columns=self.metadata_columns)
        else:
            self.metadata_df = pd.DataFrame(columns=self.metadata_columns)
    
    def setup_ui(self):
        """Create unified interface with batch mode toggle"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=5)
        ttk.Label(header_frame, text="NCBI Sequence Fetcher", font=('Helvetica', 14, 'bold')).pack(side=tk.LEFT)
        
        # Batch Mode Toggle
        ttk.Checkbutton(header_frame, text="Batch Mode", 
                       variable=self.batch_mode,
                       command=self.toggle_batch_mode).pack(side=tk.RIGHT, padx=10)
        
        # URL Input Section
        url_frame = ttk.Frame(main_frame)
        url_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(url_frame, text="NCBI URLs:").pack(anchor=tk.W)
        
        # URL Input - Will be swapped between Entry and ScrolledText
        self.url_container = ttk.Frame(url_frame)
        self.url_container.pack(fill=tk.BOTH, expand=True)
        
        # Single URL Entry
        self.url_entry = ttk.Entry(self.url_container, textvariable=self.ncbi_url)
        self.url_entry.pack(fill=tk.X, expand=True)
        self.url_entry.bind('<FocusIn>', lambda e: self.url_entry.select_range(0, tk.END))
        
        # Batch URL Text (hidden initially)
        self.batch_url_text = scrolledtext.ScrolledText(self.url_container, height=8, wrap=tk.WORD)
        
        # Format and Folder Selection
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Format Selection
        ttk.Label(control_frame, text="Format:").pack(side=tk.LEFT)
        ttk.OptionMenu(control_frame, self.report_type, "fasta", "fasta", "genbank").pack(side=tk.LEFT, padx=5)
        
        # Output Folder
        ttk.Label(control_frame, text="Save To:").pack(side=tk.LEFT, padx=(20,0))
        ttk.Entry(control_frame, textvariable=self.output_folder, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(control_frame, text="Browse", command=self.browse_folder).pack(side=tk.LEFT)
        
        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.download_btn = ttk.Button(button_frame, text="Download", command=self.start_download)
        self.download_btn.pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="Clear", command=self.clear_urls).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Import URLs", command=self.import_urls).pack(side=tk.LEFT)
        
        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, orient="horizontal",
                                          mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        # Log Area
        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status Bar
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, padx=5, pady=5)
    
    def toggle_batch_mode(self):
        """Switch between single and batch URL input"""
        if self.batch_mode.get():
            # Switch to batch mode
            current_url = self.ncbi_url.get()
            self.url_entry.pack_forget()
            self.batch_url_text.pack(fill=tk.BOTH, expand=True)
            if current_url:
                self.batch_url_text.insert(tk.END, current_url + "\n")
            self.download_btn.config(text="Download All")
        else:
            # Switch to single mode
            self.batch_url_text.pack_forget()
            self.url_entry.pack(fill=tk.X, expand=True)
            self.download_btn.config(text="Download")
    
    def browse_folder(self):
        """Select output directory"""
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)
            self.metadata_file = os.path.join(folder, "ncbi_metadata.xlsx")
            self.log("Output folder set to: " + folder)
    
    def clear_urls(self):
        """Clear URL inputs"""
        if self.batch_mode.get():
            self.batch_url_text.delete(1.0, tk.END)
        else:
            self.ncbi_url.set("")
    
    def import_urls(self):
        """Import URLs from file"""
        filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filepath:
            with open(filepath, 'r') as f:
                if self.batch_mode.get():
                    self.batch_url_text.insert(tk.END, f.read())
                else:
                    # Use first line for single mode
                    first_line = f.readline().strip()
                    self.ncbi_url.set(first_line)
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def start_download(self):
        """Start download process based on mode"""
        if self.running:
            return
            
        try:
            self.running = True
            if self.batch_mode.get():
                self.download_batch()
            else:
                self.download_single()
        except Exception as e:
            self.log(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Operation failed: {str(e)}")
        finally:
            self.running = False
    
    def download_single(self):
        """Process single URL"""
        url = self.ncbi_url.get().strip()
        if not url:
            messagebox.showwarning("Warning", "Please enter a valid NCBI URL")
            return
        
        self.process_url(url)
        messagebox.showinfo("Success", "Download completed!")
    
    def download_batch(self):
        """Process multiple URLs"""
        urls = self.get_urls_from_batch()
        if not urls:
            messagebox.showwarning("Warning", "No valid URLs found!")
            return
        
        self.log(f"\nStarting batch download of {len(urls)} URLs...")
        self.progress_var.set(0)
        self.progress_bar["maximum"] = len(urls)
        
        # Process in parallel
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self.process_url, url): url for url in urls}
            
            for i, future in enumerate(as_completed(futures), 1):
                url = futures[future]
                try:
                    result = future.result()
                    self.log(f"Completed: {url}")
                except Exception as e:
                    self.log(f"Failed {url}: {str(e)}")
                
                self.progress_var.set(i)
                self.update_status(f"Processed {i}/{len(urls)}")
                self.root.update()
        
        messagebox.showinfo("Complete", f"Finished processing {len(urls)} URLs")
    
    def get_urls_from_batch(self):
        """Extract URLs from batch text"""
        text = self.batch_url_text.get(1.0, tk.END).strip()
        return [url.strip() for url in text.split('\n') if url.strip()]
    
    def process_url(self, url):
        """Download and process a single URL"""
        self.log(f"\nProcessing: {url}")
        
        # Validate URL
        if not url.startswith(("https://www.ncbi.nlm.nih.gov/", "http://www.ncbi.nlm.nih.gov/")):
            raise ValueError("Invalid NCBI URL format")
        
        # Extract accession ID
        accession_id = url.split("/")[-1].split("?")[0]
        if not accession_id:
            raise ValueError("Could not extract accession ID")
        
        # Download data
        ext = "fasta" if self.report_type.get() == "fasta" else "gb"
        api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype={ext}&retmode=text"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/plain',
        }
        
        response = requests.get(api_url, headers=headers, timeout=15)
        response.raise_for_status()
        
        # Validate response
        if ext == "fasta" and not response.text.startswith('>'):
            raise ValueError("Invalid FASTA format")
        elif ext == "gb" and not response.text.startswith('LOCUS'):
            raise ValueError("Invalid GenBank format")
        
        # Extract metadata
        metadata = self.extract_metadata(accession_id)
        
        # Save files
        filename = self.save_data(response.text, metadata, ext)
        self.save_metadata(metadata, filename)
        
        return filename
    
    def extract_metadata(self, accession_id):
        """Extract metadata from GenBank format"""
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
            # Get GenBank format for metadata
            api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype=gb&retmode=text"
            response = requests.get(api_url, headers={
                'User-Agent': 'Mozilla/5.0...'
            }, timeout=10)
            response.raise_for_status()
            gb_data = response.text

            # Parse metadata
            lines = gb_data.split('\n')
            for i, line in enumerate(lines):
                try:
                    if line.startswith('VERSION'):
                        metadata['Version'] = line.split()[1].strip()
                    elif line.startswith('DEFINITION'):
                        metadata['Definition'] = line[12:].strip()
                    elif line.startswith('  ORGANISM'):
                        metadata['Organism'] = line[12:].strip()
                        # Get taxonomy
                        taxonomy_lines = []
                        j = i + 1
                        while j < len(lines) and lines[j].startswith(' ' * 10):
                            taxonomy_lines.append(lines[j].strip())
                            j += 1
                        if taxonomy_lines:
                            metadata['Taxonomy'] = '; '.join(taxonomy_lines).strip('; ')
                    elif '/strain=' in line:
                        metadata['Strain'] = self.extract_value(line, 'strain')
                    elif '/country=' in line or '/geo_loc_name=' in line:
                        if '/country=' in line:
                            metadata['Country'] = self.extract_value(line, 'country')
                        else:
                            metadata['Country'] = self.extract_value(line, 'geo_loc_name')
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
                except:
                    continue

        except Exception as e:
            self.log(f"Metadata warning: {str(e)}")

        return metadata

    def extract_value(self, line, field_name):
        """Extract value from line"""
        patterns = [f'/{field_name}="', f'/{field_name}=']
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
        """Save sequence data to file"""
        # Generate filename components
        components = [
            metadata.get('Organism', 'unknown').replace(' ', '_'),
            metadata.get('Strain', 'unknown'),
            metadata.get('Accession', ''),
            metadata.get('Length', ''),
            metadata.get('Product', '')
        ]
        
        # Add description
        desc = []
        definition = metadata.get('Definition', '').lower()
        if 'partial' in definition:
            desc.append('partial')
        if 'complete' in definition:
            desc.append('complete')
        if 'cds' in definition:
            desc.append('cds')
        elif 'gene' in definition:
            desc.append('gene')
        elif 'genome' in definition:
            desc.append('genome')
        
        if desc:
            components.append('_'.join(desc))
        
        # Create filename
        filename = '_'.join(filter(None, components)) + f".{ext}"
        filename = "".join(c if c.isalnum() or c in ('_', '-', '.') else '_' for c in filename)
        
        # Save file
        filepath = os.path.join(self.output_folder.get(), filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(data)
        
        self.log(f"Saved: {filename}")
        return filename
    
    def save_metadata(self, metadata, filename):
        """Update metadata Excel"""
        try:
            metadata['Filename'] = filename
            new_row = pd.DataFrame([metadata])
            self.metadata_df = pd.concat([self.metadata_df, new_row], ignore_index=True)
            self.metadata_df.to_excel(self.metadata_file, index=False)
        except Exception as e:
            self.log(f"Metadata error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NCBISequenceFetcher(root)
    root.mainloop()