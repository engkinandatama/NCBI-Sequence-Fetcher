import os
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime

class NCBIScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("NCBI Data Scraper")
        self.root.geometry("700x500")
        
        # Variables
        self.output_folder = tk.StringVar(value=os.path.expanduser("~"))
        self.ncbi_url = tk.StringVar()
        self.report_type = tk.StringVar(value="fasta")  # Default to FASTA
        
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
        """Create the user interface"""
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # URL Input
        ttk.Label(main_frame, text="NCBI URL:").grid(row=0, column=0, sticky=tk.W)
        self.url_entry = ttk.Entry(main_frame, textvariable=self.ncbi_url, width=60)
        self.url_entry.grid(row=1, column=0, columnspan=2, pady=5, sticky=tk.EW)
        
        # Bind focus in event to select all text
        self.url_entry.bind('<FocusIn>', lambda e: self.url_entry.select_range(0, tk.END))
        
        # Report Type - Changed to OptionMenu for non-editable dropdown
        ttk.Label(main_frame, text="Format:").grid(row=2, column=0, sticky=tk.W)
        report_menu = ttk.OptionMenu(main_frame, self.report_type, "fasta", "fasta", "genbank")
        report_menu.grid(row=2, column=1, sticky=tk.W, padx=5)
        
        # Output Folder
        ttk.Label(main_frame, text="Save To:").grid(row=3, column=0, sticky=tk.W)
        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=4, column=0, columnspan=2, sticky=tk.EW)
        
        ttk.Entry(folder_frame, textvariable=self.output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(folder_frame, text="Browse", command=self.browse_folder).pack(side=tk.LEFT, padx=5)
        
        # Download Button
        ttk.Button(main_frame, text="Download", command=self.download_data).grid(row=5, column=0, pady=10)
        
        # Log Area
        self.log_text = tk.Text(main_frame, height=12, wrap=tk.WORD)
        self.log_text.grid(row=6, column=0, columnspan=2, sticky=tk.NSEW)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=6, column=2, sticky=tk.NS)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Bind Enter key
        self.root.bind('<Return>', lambda e: self.download_data())
    
    def browse_folder(self):
        """Select output directory"""
        folder = filedialog.askdirectory(initialdir=self.output_folder.get())
        if folder:
            self.output_folder.set(folder)
            self.metadata_file = os.path.join(folder, "ncbi_metadata.xlsx")
            self.log(f"Output folder set to: {folder}")
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def download_data(self):
        """Main download function"""
        url = self.ncbi_url.get().strip()
        if not url:
            messagebox.showerror("Error", "Please enter a valid NCBI URL")
            return
        
        try:
            # Validate URL
            if not url.startswith(("https://www.ncbi.nlm.nih.gov/", "http://www.ncbi.nlm.nih.gov/")):
                messagebox.showerror("Error", "URL must be from NCBI (e.g. https://www.ncbi.nlm.nih.gov/nuccore/JN188370.1)")
                return
            
            # Extract accession ID
            accession_id = url.split("/")[-1].split("?")[0]
            if not accession_id:
                raise ValueError("Could not extract accession ID from URL")
            
            # Generate API URL based on report type
            if self.report_type.get() == "fasta":
                api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype=fasta&retmode=text"
            else:  # genbank
                api_url = f"https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=nuccore&id={accession_id}&rettype=gb&retmode=text"
            
            self.log(f"Fetching data from NCBI API...")
            
            # Set headers to mimic browser
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                'Accept': 'text/plain',
            }
            
            # Download data
            response = requests.get(api_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            # Validate response
            if self.report_type.get() == "fasta" and not response.text.startswith('>'):
                raise ValueError("Invalid FASTA format received from NCBI")
            elif self.report_type.get() == "genbank" and not response.text.startswith('LOCUS'):
                raise ValueError("Invalid GenBank format received from NCBI")
            
            # Extract metadata from GenBank format (even if we're downloading FASTA)
            metadata = self.extract_metadata(accession_id)
            
            # Save data with proper filename based on metadata
            filename = self.save_data(response.text, metadata)
            
            # Save metadata
            self.save_metadata(metadata, filename)
            
            messagebox.showinfo("Success", "Data downloaded successfully!")
            
        except Exception as e:
            self.log(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Download failed: {str(e)}")
    
    def extract_metadata(self, accession_id):
        """Extract metadata from GenBank format with improved taxonomy and country handling"""
        # Initialize metadata with default values
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

            # Parse GenBank data line by line
            lines = gb_data.split('\n')
            for i, line in enumerate(lines):
                try:
                    if line.startswith('VERSION'):
                        metadata['Version'] = line.split()[1].strip() if len(line.split()) > 1 else 'NA'
                    elif line.startswith('DEFINITION'):
                        metadata['Definition'] = line[12:].strip() if len(line) > 12 else 'NA'
                    elif line.startswith('  ORGANISM'):
                        # Extract organism name
                        metadata['Organism'] = line[12:].strip() if len(line) > 12 else 'NA'
                        
                        # Extract full taxonomy from subsequent lines
                        taxonomy_lines = []
                        j = i + 1
                        while j < len(lines) and lines[j].startswith(' ' * 10):  # Taxonomy lines are indented
                            taxonomy_lines.append(lines[j].strip())
                            j += 1
                        
                        if taxonomy_lines:
                            # Join taxonomy lines and clean up
                            full_taxonomy = '; '.join(taxonomy_lines)
                            # Remove any trailing semicolons or spaces
                            metadata['Taxonomy'] = full_taxonomy.strip('; ')

                    elif '/strain=' in line:
                        metadata['Strain'] = self.extract_quoted_value(line, 'strain')
                    elif '/country=' in line or '/geo_loc_name=' in line:
                        # Check both possible field names for country
                        if '/country=' in line:
                            metadata['Country'] = self.extract_quoted_value(line, 'country')
                        elif '/geo_loc_name=' in line:
                            metadata['Country'] = self.extract_quoted_value(line, 'geo_loc_name')
                    elif '/collection_date=' in line:
                        metadata['Collection_Date'] = self.extract_quoted_value(line, 'collection_date')
                    elif '/collected_by=' in line:
                        metadata['Collected_By'] = self.extract_quoted_value(line, 'collected_by')
                    elif '/isolation_source=' in line:
                        metadata['Isolation_Source'] = self.extract_quoted_value(line, 'isolation_source')
                    elif '/product=' in line:
                        metadata['Product'] = self.extract_quoted_value(line, 'product')
                    elif line.startswith('LOCUS'):
                        parts = line.split()
                        if len(parts) >= 3:
                            metadata['Length'] = parts[2] + 'bp'
                except Exception as field_error:
                    self.log(f"Warning: Error parsing line '{line}': {str(field_error)}")
                    continue

        except requests.exceptions.RequestException as e:
            self.log(f"Warning: Failed to fetch GenBank data for metadata: {str(e)}")
        except Exception as e:
            self.log(f"Warning: Unexpected error during metadata extraction: {str(e)}")

        return metadata

    def extract_quoted_value(self, line, field_name):
        """Helper method to safely extract quoted values from GenBank lines"""
        try:
            # Find the field pattern (supports both quoted and unquoted values)
            pattern1 = f'/{field_name}="'
            pattern2 = f'/{field_name}='
            
            start_idx = line.find(pattern1)
            if start_idx != -1:
                # Quoted value
                start_idx += len(pattern1)
                end_idx = line.find('"', start_idx)
                if end_idx == -1:
                    return 'NA'
                return line[start_idx:end_idx].strip()
            else:
                # Unquoted value
                start_idx = line.find(pattern2)
                if start_idx == -1:
                    return 'NA'
                start_idx += len(pattern2)
                end_idx = len(line)
                # Find end of value (either space or end of line)
                space_idx = line.find(' ', start_idx)
                if space_idx != -1 and space_idx > start_idx:
                    end_idx = space_idx
                semicolon_idx = line.find(';', start_idx)
                if semicolon_idx != -1 and semicolon_idx > start_idx:
                    end_idx = min(end_idx, semicolon_idx)
                return line[start_idx:end_idx].strip()
        except:
            return 'NA'
    
    def save_data(self, data, metadata):
        """Save the downloaded data to file with proper naming"""
        # Determine file extension
        ext = ".fasta" if self.report_type.get() == "fasta" else ".gb"
        
        # Generate filename based on metadata
        organism = metadata.get('Organism', 'unknown_organism').replace(' ', '_')
        strain = metadata.get('Strain', 'unknown_strain').replace(' ', '_')
        accession = metadata.get('Accession', 'no_accession')
        length = metadata.get('Length', 'unknown_length')
        product = metadata.get('Product', 'unknown_product')
        definition = metadata.get('Definition', '')
        
        # Extract description from DEFINITION (e.g., "partial cds")
        description = ''
        if definition:
            # Look for common patterns in definition
            if 'partial' in definition.lower():
                description = 'partial'
            elif 'complete' in definition.lower():
                description = 'complete'
            if 'cds' in definition.lower():
                description += '_cds' if description else 'cds'
            elif 'gene' in definition.lower():
                description += '_gene' if description else 'gene'
            elif 'genome' in definition.lower():
                description += '_genome' if description else 'genome'
        
        # Construct filename
        filename_parts = []
        if organism: filename_parts.append(organism)
        if strain: filename_parts.append(strain)
        if accession: filename_parts.append(accession)
        if length: filename_parts.append(length)
        if product: filename_parts.append(product)
        if description: filename_parts.append(description)
        
        filename = '_'.join(filename_parts) + ext
        
        # Clean filename
        filename = "".join(c for c in filename if c.isalnum() or c in ('_', '-', '.'))
        filepath = os.path.join(self.output_folder.get(), filename)
        
        # Save file
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(data)
        
        self.log(f"File saved as: {filepath}")
        return filename
    
    def save_metadata(self, metadata, filename):
        """Save metadata to Excel file"""
        try:
            metadata['Filename'] = filename
            
            # Add to DataFrame and save
            self.metadata_df = pd.concat([self.metadata_df, pd.DataFrame([metadata])], ignore_index=True)
            self.metadata_df.to_excel(self.metadata_file, index=False)
            self.log("Metadata saved successfully")
            
        except Exception as e:
            self.log(f"Warning: Metadata save failed - {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = NCBIScraperApp(root)
    root.mainloop()