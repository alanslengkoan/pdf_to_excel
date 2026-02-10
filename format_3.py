# ============================================================
# KONVERSI REKENING KORAN BRI - COORDINATE-BASED EXTRACTION
# Menggunakan posisi koordinat untuk extract data yang akurat
# ============================================================

!pip install pdfplumber pandas openpyxl -q
print("✅ Libraries installed!\n")

from google.colab import files
print("📤 Upload PDF rekening koran BRI:")
uploaded = files.upload()
pdf_file = list(uploaded.keys())[0]
print(f"✅ {pdf_file}\n")

import pdfplumber
import pandas as pd
import re

print("🔄 Extracting data...\n")

all_rows = []
info = {}

# ============================================================
# FUNGSI KONVERSI FORMAT ANGKA (US → ID)
# ============================================================
def convert_to_indonesian_format(value):
    """Konversi format angka dari 2,662,608.00 menjadi 2.662.608,00"""
    if pd.isna(value) or value == '' or value == '0.00':
        return ''
    
    value_str = str(value).strip()
    value_str = re.sub(r'[^\d,.-]', '', value_str)
    
    if not value_str or value_str == '-' or value_str == '0.00':
        return ''
    
    try:
        if ',' in value_str and '.' in value_str:
            value_str = value_str.replace(',', '')
            number = float(value_str)
        elif ',' in value_str:
            value_str = value_str.replace(',', '')
            number = float(value_str)
        elif '.' in value_str:
            number = float(value_str)
        else:
            number = float(value_str)
        
        formatted = f"{number:,.2f}"
        formatted = formatted.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
        return formatted
        
    except (ValueError, AttributeError):
        return value_str

# ============================================================
# FUNGSI PEMBERSIH TEXT
# ============================================================
def clean_text(text):
    """Bersihkan karakter aneh dari text"""
    if pd.isna(text) or text == '':
        return ''
    
    text = str(text)
    text = re.sub(r'[^\x20-\x7E\n]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

with pdfplumber.open(pdf_file) as pdf:
    total_pages = len(pdf.pages)
    print(f"📄 Total pages: {total_pages}\n")
    
    # Get header info
    first_text = pdf.pages[0].extract_text()
    lines = first_text.split('\n')
    
    for i, line in enumerate(lines):
        if 'Kepada Yth' in line or 'To :' in line:
            if i + 1 < len(lines):
                info['Nama'] = lines[i + 1].strip()
        
        if 'No. Rekening' in line or 'Account No' in line:
            match = re.search(r':?\s*(\d+)', line)
            if match:
                info['No_Rekening'] = match.group(1)
        
        if 'Periode Transaksi' in line or 'Transaction Period' in line:
            match = re.search(r':?\s*(\d{2}/\d{2}/\d{2,4}\s*-\s*\d{2}/\d{2}/\d{2,4})', line)
            if match:
                info['Periode'] = match.group(1)
        
        if 'Nama Produk' in line or 'Product Name' in line:
            parts = line.split(':')
            if len(parts) > 1:
                info['Produk'] = parts[1].strip()
    
    print("🔍 Analyzing table structure from first page...\n")
    
    # Analyze column positions from first page
    first_page = pdf.pages[0]
    words = first_page.extract_words()
    
    # Cari header table untuk tentukan posisi kolom
    header_y = None
    col_positions = {}
    
    for word in words:
        if 'Tanggal' in word['text'] and 'Transaksi' in word['text']:
            header_y = word['top']
        elif word['text'] == 'Debet' or word['text'] == 'Debit':
            col_positions['debet_x'] = word['x0']
        elif word['text'] == 'Kredit' or word['text'] == 'Credit':
            col_positions['kredit_x'] = word['x0']
        elif word['text'] == 'Saldo' or word['text'] == 'Balance':
            col_positions['saldo_x'] = word['x0']
        elif word['text'] == 'Teller' or word['text'] == 'User':
            col_positions['teller_x'] = word['x0']
    
    print(f"   Column positions detected:")
    for key, val in col_positions.items():
        print(f"   - {key}: {val:.1f}")
    
    print("\n🔍 Extracting transactions using table method...\n")
    
    # Process each page
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"   Processing page {page_num}/{total_pages}...", end='\r')
        
        # Use text-based table extraction
        tables = page.extract_tables({
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
        })
        
        if tables:
            for table in tables:
                if not table:
                    continue
                
                for row in table:
                    if not row or len(row) < 6:
                        continue
                    
                    # Gabungkan semua cell dalam row
                    row_text = ' '.join([str(c) if c else '' for c in row])
                    
                    # Skip header
                    if 'Tanggal Transaksi' in row_text or 'Transaction Date' in row_text:
                        continue
                    if 'Debet' in row_text and 'Kredit' in row_text and 'Saldo' in row_text:
                        continue
                    
                    # Cari tanggal dalam row
                    date_match = re.search(r'(\d{2}/\d{2}/\d{2,4}(?:\s+\d{2}:\d{2}:\d{2})?)', row_text)
                    if not date_match:
                        continue
                    
                    tanggal = date_match.group(1)
                    
                    # Cari semua angka (amounts) dalam format: 1,234.56 atau 1234.56
                    amounts = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d{2})', row_text)
                    
                    if len(amounts) < 1:
                        continue
                    
                    # Ambil 3 angka terakhir sebagai Debet, Kredit, Saldo
                    # Atau 1 angka terakhir sebagai Saldo saja
                    if len(amounts) >= 3:
                        debet = amounts[-3]
                        kredit = amounts[-2]
                        saldo = amounts[-1]
                    elif len(amounts) == 2:
                        debet = '0.00'
                        kredit = amounts[-2]
                        saldo = amounts[-1]
                    else:
                        debet = '0.00'
                        kredit = '0.00'
                        saldo = amounts[-1]
                    
                    # Extract Uraian - text antara tanggal dan angka pertama
                    # Hapus tanggal
                    rest = row_text.replace(tanggal, '', 1).strip()
                    
                    # Hapus semua amounts
                    for amt in amounts:
                        rest = rest.replace(amt, '', 1)
                    
                    # Clean up
                    rest = re.sub(r'\s+', ' ', rest).strip()
                    
                    # Split untuk dapat Uraian dan Teller
                    parts = rest.split()
                    
                    # Teller biasanya kode singkat di akhir atau all-caps
                    teller = ''
                    uraian_parts = []
                    
                    for part in reversed(parts):
                        if part.isupper() and len(part) >= 4:
                            teller = part
                            break
                    
                    if teller:
                        uraian_parts = [p for p in parts if p != teller]
                    else:
                        # Ambil yang terakhir sebagai teller
                        if len(parts) > 1:
                            teller = parts[-1]
                            uraian_parts = parts[:-1]
                        else:
                            uraian_parts = parts
                    
                    uraian = ' '.join(uraian_parts).strip()
                    
                    # Clean debet/kredit
                    if debet == '0.00':
                        debet = ''
                    if kredit == '0.00':
                        kredit = ''
                    
                    all_rows.append([tanggal, uraian, teller, debet, kredit, saldo])

print(f"\n\n✅ Extracted: {len(all_rows):,} rows\n")

if len(all_rows) == 0:
    print("❌ Tidak ada data yang terekstrak!")
    print("\nDEBUG: Tampilkan raw text untuk analisa:")
    with pdfplumber.open(pdf_file) as pdf:
        text = pdf.pages[0].extract_text()
        lines = text.split('\n')
        print("\nBaris 20-40 dari halaman 1:")
        for i, line in enumerate(lines[20:40]):
            if re.search(r'\d{2}/\d{2}/\d{2}', line):
                print(f"   >>> {i+20}: {line}")
            else:
                print(f"       {i+20}: {line}")
else:
    # Create DataFrame
    df = pd.DataFrame(all_rows, columns=[
        'Tanggal Transaksi', 'Uraian Transaksi', 'Teller', 
        'Debet', 'Kredit', 'Saldo'
    ])
    
    df = df.replace('', pd.NA)
    df = df.dropna(how='all')
    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    
    print(f"🔄 Cleaning and formatting data...\n")
    
    df['Uraian Transaksi'] = df['Uraian Transaksi'].fillna('').apply(clean_text)
    df['Teller'] = df['Teller'].fillna('').apply(clean_text)
    
    df['Debet'] = df['Debet'].fillna('').apply(convert_to_indonesian_format)
    df['Kredit'] = df['Kredit'].fillna('').apply(convert_to_indonesian_format)
    df['Saldo'] = df['Saldo'].fillna('').apply(convert_to_indonesian_format)
    
    # Hitung total
    total_debet = df['Debet'].apply(lambda x: float(x.replace('.', '').replace(',', '.')) if x else 0).sum()
    total_kredit = df['Kredit'].apply(lambda x: float(x.replace('.', '').replace(',', '.')) if x else 0).sum()
    
    info['Total_Debet'] = f"{total_debet:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
    info['Total_Kredit'] = f"{total_kredit:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
    
    print(f"📊 Final data: {len(df):,} rows\n")
    
    # Save to Excel
    output = 'rekening_koran_bri.xlsx'
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        info_data = [
            ['Nama Rekening', info.get('Nama', '')],
            ['No. Rekening', info.get('No_Rekening', '')],
            ['Periode', info.get('Periode', '')],
            ['Produk', info.get('Produk', '')],
            ['Total Debet', info.get('Total_Debet', '')],
            ['Total Kredit', info.get('Total_Kredit', '')],
            ['Total Transaksi', f"{len(df):,}"],
        ]
        pd.DataFrame(info_data).to_excel(writer, sheet_name='Info', index=False, header=False)
        df.to_excel(writer, sheet_name='Transaksi', index=False)
    
    print("✅ Excel saved\n")
    
    print("="*140)
    print("INFORMASI REKENING")
    print("="*140)
    for key, value in info.items():
        print(f"{key:20s}: {value}")
    
    print("\n" + "="*140)
    print("PREVIEW TRANSAKSI (15 Pertama)")
    print("="*140)
    pd.set_option('display.max_colwidth', 70)
    pd.set_option('display.width', None)
    print(df.head(15).to_string(index=False))
    
    print(f"\n\n📊 Summary:")
    print(f"   Total Transaksi : {len(df):,}")
    print(f"   Transaksi Debet : {len(df[df['Debet'] != '']):,}")
    print(f"   Transaksi Kredit: {len(df[df['Kredit'] != '']):,}")
    
    print(f"\n{'='*140}")
    
    files.download(output)
    df.to_csv('rekening_koran_bri.csv', index=False, encoding='utf-8-sig')
    files.download('rekening_koran_bri.csv')
    print(f"\n🎉 Done! {len(df):,} transaksi berhasil dikonversi")
    print("📥 File: rekening_koran_bri.xlsx & rekening_koran_bri.csv")