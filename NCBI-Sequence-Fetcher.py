import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

class NCBIScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NCBI Batch Downloader")
        self.root.geometry("800x600")
        
        # Variables
        self.output_folder = tk.StringVar(value=os.path.expanduser("~"))
        self.ncbi_url = tk.StringVar()
        self.report_type = tk.StringVar(value="fasta")
        self.running = False  # Flag for batch processing
        
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
        """Create the user interface with tabs"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tab System
        tab_control = ttk.Notebook(main_frame)
        
        # Single Download Tab
        single_tab = ttk.Frame(tab_control)
        self.setup_single_tab(single_tab)
        
        # Batch Download Tab
        batch_tab = ttk.Frame(tab_control)
        self.setup_batch_tab(batch_tab)
        
        tab_control.add(single_tab, text="Single Download")
        tab_control.add(batch_tab, text="Batch Download")
        tab_control.pack(expand=1, fill="both")
        
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, padx=5, pady=5)
    
    def setup_single_tab(self, parent):
        """Setup single download tab"""
        # URL Input
        ttk.Label(parent, text="NCBI URL:").pack(anchor=tk.W, pady=(5,0))
        self.url_entry = ttk.Entry(parent, textvariable=self.ncbi_url, width=80)
        self.url_entry.pack(fill=tk.X, padx=5, pady=5)
        self.url_entry.bind('<FocusIn>', lambda e: self.url_entry.select_range(0, tk.END))
        
        # Format Selection
        format_frame = ttk.Frame(parent)
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(format_frame, text="Format:").pack(side=tk.LEFT)
        ttk.OptionMenu(format_frame, self.report_type, "fasta", "fasta", "genbank").pack(side=tk.LEFT, padx=5)
        
        # Output Folder
        folder_frame = ttk.Frame(parent)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(folder_frame, text="Save To:").pack(side=tk.LEFT)
        ttk.Entry(folder_frame, textvariable=self.output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side=tk.LEFT)
        
        # Download Button
        ttk.Button(parent, text="Download", command=self.download_single).pack(pady=10)
        
        # Log Area
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def setup_batch_tab(self, parent):
        """Setup batch download tab"""
        # URL Input Area
        ttk.Label(parent, text="Enter NCBI URLs (one per line):").pack(anchor=tk.W, pady=(5,0))
        self.batch_url_text = scrolledtext.ScrolledText(parent, height=12, wrap=tk.WORD)
        self.batch_url_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Control Buttons
        control_frame = ttk.Frame(parent)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(control_frame, text="Import from File", 
                  command=self.import_urls_from_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="Clear List",
                  command=self.clear_batch_urls).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="Download All", 
                  command=self.download_batch).pack(side=tk.RIGHT, padx=2)
        
        # Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(parent, orient="horizontal", 
                                          length=200, mode="determinate",
                                          variable=self.progress_var)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        # Batch Log
        self.batch_log = scrolledtext.ScrolledText(parent, height=8, wrap=tk.WORD)
        self.batch_log.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    def browse_folder(self):
        """Select output directory"""
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)
            self.metadata_file = os.path.join(folder, "ncbi_metadata.xlsx")
            self.log("Output folder set to: " + folder)
    
    def log(self, message, batch=False):
        """Add message to log"""
        if batch:
            self.batch_log.insert(tk.END, message + "\n")
            self.batch_log.see(tk.END)
        else:
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def download_single(self):
        """Handle single download"""
        if self.running:
            return
            
        url = self.ncbi_url.get().strip()
        if not url:
            messagebox.showwarning("Warning", "Please enter a valid NCBI URL")
            return
        
        try:
            self.running = True
            self.process_url(url)
        except Exception as e:
            self.log(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Download failed: {str(e)}")
        finally:
            self.running = False
    
    def download_batch(self):
        """Handle batch download"""
        if self.running:
            return
            
        urls = self.get_urls_from_text()
        if not urls:
            messagebox.showwarning("Warning", "No URLs entered!")
            return
        
        self.running = True
        self.batch_log.delete(1.0, tk.END)
        self.progress_var.set(0)
        self.progress_bar["maximum"] = len(urls)
        
        # Use thread pool for concurrent downloads
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = {executor.submit(self.process_url, url): url for url in urls}
            
            for i, future in enumerate(as_completed(futures), 1):
                url = futures[future]
                try:
                    result = future.result()
                    self.log(f"Completed: {url}", batch=True)
                except Exception as e:
                    self.log(f"Error processing {url}: {str(e)}", batch=True)
                
                self.progress_var.set(i)
                self.update_status(f"Processing {i}/{len(urls)}")
                self.root.update()
        
        self.running = False
        messagebox.showinfo("Complete", f"Finished processing {len(urls)} URLs")
    
    def get_urls_from_text(self):
        """Extract URLs from text widget"""
        text = self.batch_url_text.get(1.0, tk.END).strip()
        return [url.strip() for url in text.split('\n') if url.strip()]
    
    def import_urls_from_file(self):
        """Import URLs from text file"""
        filepath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filepath:
            with open(filepath, 'r') as f:
                self.batch_url_text.insert(tk.END, f.read())
    
    def clear_batch_urls(self):
        """Clear the URL list"""
        self.batch_url_text.delete(1.0, tk.END)
    
    def process_url(self, url):
        """Process a single URL (download + save)"""
        self.log(f"\nProcessing: {url}")
        
        # Validate URL
        if not url.startswith(("https://www.ncbi.nlm.nih.gov/", "http://www.ncbi.nlm.nih.gov/")):
            raise ValueError("Invalid NCBI URL format")
        
        # Extract accession ID
        accession_id = url.split("/")[-1].split("?")[0]
        if not accession_id:
            raise ValueError("Could not extract accession ID from URL")
        
        # Generate API URL
        if self.report_type.get() == "fasta":
            api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype=fasta&retmode=text"
        else:  # genbank
            api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype=gb&retmode=text"
        
        # Download data
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/plain',
        }
        
        response = requests.get(api_url, headers=headers, timeout=15)
        response.raise_for_status()
        
        # Validate response
        if self.report_type.get() == "fasta" and not response.text.startswith('>'):
            raise ValueError("Invalid FASTA format received")
        elif self.report_type.get() == "genbank" and not response.text.startswith('LOCUS'):
            raise ValueError("Invalid GenBank format received")
        
        # Extract metadata
        metadata = self.extract_metadata(accession_id)
        
        # Save files
        filename = self.save_data(response.text, metadata)
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

            # Parse GenBank data
            lines = gb_data.split('\n')
            for i, line in enumerate(lines):
                try:
                    if line.startswith('VERSION'):
                        metadata['Version'] = line.split()[1].strip()
                    elif line.startswith('DEFINITION'):
                        metadata['Definition'] = line[12:].strip()
                    elif line.startswith('  ORGANISM'):
                        metadata['Organism'] = line[12:].strip()
                        # Extract taxonomy
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
                except Exception as e:
                    continue

        except Exception as e:
            self.log(f"Metadata extraction warning: {str(e)}")

        return metadata

    def extract_value(self, line, field_name):
        """Extract value from GenBank line"""
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
    
    def save_data(self, data, metadata):
        """Save sequence data to file"""
        ext = ".fasta" if self.report_type.get() == "fasta" else ".gb"
        
        # Generate filename components
        components = [
            metadata.get('Organism', 'unknown').replace(' ', '_'),
            metadata.get('Strain', 'unknown'),
            metadata.get('Accession', ''),
            metadata.get('Length', ''),
            metadata.get('Product', '')
        ]
        
        # Add description from definition
        definition = metadata.get('Definition', '').lower()
        desc = []
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
        
        # Filter and join components
        filename = '_'.join(filter(None, components)) + ext
        filename = "".join(c if c.isalnum() or c in ('_', '-', '.') else '_' for c in filename)
        
        # Save file
        filepath = os.path.join(self.output_folder.get(), filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(data)
        
        self.log(f"Saved: {filename}")
        return filename
    
    def save_metadata(self, metadata, filename):
        """Update metadata Excel file"""
        try:
            metadata['Filename'] = filename
            new_row = pd.DataFrame([metadata])
            self.metadata_df = pd.concat([self.metadata_df, new_row], ignore_index=True)
            self.metadata_df.to_excel(self.metadata_file, index=False)
        except Exception as e:
            self.log(f"Metadata save error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NCBIScraperApp(root)
    root.mainloop()