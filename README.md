# 🧬 NCBI Sequence Fetcher

A lightweight and user-friendly desktop application for downloading nucleotide sequences and extracting biological metadata directly from NCBI.  
Built with Python and Tkinter, designed for researchers, students, and bioinformaticians.

---

### ⚠️ **Note**:
> **This repository is for personal and educational use only.**  
> It is not currently open for external collaboration or contribution.  
> Please use responsibly and cite NCBI appropriately when using downloaded data.

---

## 📸 GUI Preview

Coming soon — stay tuned!

---

## ✨ Features

- 🔗 **Direct URL Input**: Download GenBank/FASTA files from any valid NCBI nuccore link.
- 📁 **FASTA / GenBank Format Support**: Choose your preferred sequence format.
- 🧬 **Automatic Metadata Extraction**:
  - Accession, Organism, Strain, Taxonomy, Country, Collection Date, Length, etc.
- 📄 **Excel Export**:
  - All metadata saved in a clean Excel file (`ncbi_metadata.xlsx`)
- 🏷️ **Smart File Naming**:
  - Files saved with informative names: `Organism_Strain_Accession_Length_Feature.fasta`
- 🖥️ **GUI Based**:
  - No command line needed; simple Tkinter-based interface.

---

## 🛠️ Technologies Used

- **Language**: Python 3.8+
- **Libraries**:
  - `requests`
  - `pandas`
  - `openpyxl`
  - `tkinter` (built-in)

---

## 🚀 Installation & Usage

### 🔧 Step 1: Clone the repository

```bash
git clone https://github.com/engkinandatama/NCBI-Sequence-Fetcher.git
```
```
cd ncbi-data-scraper
```
### 📦 Step 2: Install dependencies
```
pip install -r requirements.txt
```
### ▶️ Step 3: Run the app
```
python ncbi_scraper.py
```

---

## 🧪 How It Works

1. **Paste a valid NCBI URL**  
   Example:  https://www.ncbi.nlm.nih.gov/nuccore/JN188370.1

2. **Choose the format**  
`fasta` or `genbank`

3. **Select a destination folder**  
Where the sequence file and Excel metadata will be saved

4. **Click `Download`**

---

### 🔄 Behind the Scenes:

The app will:

- 🔍 **Fetch** the nucleotide data directly from NCBI
- 💾 **Save** the sequence locally as: `.fasta` if FASTA format is selected, `.gb` (GenBank) if GenBank format is selected
- 🧬 **Parse and extract metadata**: Accession, Organism, Strain, Country, Date, and more
- 📊 **Append** the metadata into an Excel file: `ncbi_metadata.xlsx`


---

## 📁 Output Example

Setelah proses selesai, file output akan tersimpan seperti berikut:
```
📂 Output_Folder/
├── Escherichia_coli_K12_JN188370.1_4500bp_partial_cds.fasta
└── ncbi_metadata.xlsx
```
- **FASTA / GenBank File**: Berisi urutan nukleotida yang diunduh dari NCBI.
- **ncbi_metadata.xlsx**: File Excel yang berisi metadata terstruktur dari setiap entri GenBank yang diunduh.


---

## 📜 License

This project is licensed under the **MIT License**.

> **Disclaimer**:  
> This tool is developed solely for **academic and personal research purposes**.  
> Commercial use, bulk data scraping, or redistribution of NCBI content may **violate NCBI's usage policies** and is **strongly discouraged**.  
>  
> The developer **does not take any responsibility** for misuse, legal issues, or policy violations resulting from the use of this tool.  
> **Users are fully responsible** for ensuring their usage complies with relevant terms, laws, and guidelines.  
>  
> See [NCBI's policies](https://www.ncbi.nlm.nih.gov/home/about/policies/) for more information.
