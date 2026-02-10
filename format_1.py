# ============================================================
# KONVERSI REKENING KORAN - METODE ALTERNATIF
# Menggunakan extract_text() untuk backup jika extract_tables() gagal
# ============================================================

!pip install pdfplumber pandas openpyxl -q
print("✅ Libraries installed!\n")

from google.colab import files
print("📤 Upload PDF:")
uploaded = files.upload()
pdf_file = list(uploaded.keys())[0]
print(f"✅ {pdf_file}\n")

import pdfplumber
import pandas as pd
import re

print("🔄 Extracting data...")

all_rows = []
info = {}

with pdfplumber.open(pdf_file) as pdf:
    total_pages = len(pdf.pages)
    print(f"📄 Total pages: {total_pages}\n")
    
    # Get header info
    first_text = pdf.pages[0].extract_text()
    for line in first_text.split('\n'):
        if 'Periode' in line and ':' in line: 
            info['Periode'] = line.split(':')[-1].strip()
        if 'Nama Tercetak' in line and ':' in line: 
            info['Nama'] = line.split(':')[-1].strip()
        if 'Nomor Rekening' in line and ':' in line: 
            info['No_Rek'] = line.split(':')[-1].strip()
    
    # Process each page
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"Processing page {page_num}/{total_pages}...", end='\r')
        
        # METHOD 1: Try extract_tables first
        tables = page.extract_tables()
        page_data = []
        
        if tables:
            for table in tables:
                if not table:
                    continue
                    
                for row in table:
                    if not row or len(row) < 9:
                        continue
                    
                    # Skip headers
                    row_text = ' '.join([str(c) if c else '' for c in row]).lower()
                    if ('no.' in row_text and 'debit' in row_text) or 'tgl dan waktu' in row_text:
                        continue
                    
                    # Check if valid data row
                    if row[0] or row[1]:  # Has number or date
                        page_data.append([str(c).strip() if c else '' for c in row[:9]])
        
        # METHOD 2: If no data from tables, try text extraction
        if len(page_data) == 0:
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line in lines:
                    # Skip empty lines and headers
                    if not line.strip():
                        continue
                    if 'Debit' in line and 'Kredit' in line:
                        continue
                    
                    # Try to parse line as table row
                    # Format: No | Date | Ref | Desc | Code | D/K | Debit | Kredit | Saldo
                    parts = re.split(r'\s{2,}', line.strip())  # Split by 2+ spaces
                    
                    if len(parts) >= 9:
                        # Check if first column is number
                        if parts[0].replace('.','').isdigit():
                            page_data.append(parts[:9])
        
        all_rows.extend(page_data)

print(f"\n✅ Extracted: {len(all_rows):,} rows\n")

# Create DataFrame
df = pd.DataFrame(all_rows, columns=[
    'No', 'Tgl dan Waktu', 'No Referensi', 'Deskripsi', 
    'Kode', 'D/K', 'Debit', 'Kredit', 'Saldo'
])

# Clean
df = df.replace('', pd.NA)
df = df.dropna(how='all')
df = df.drop_duplicates()
df = df.reset_index(drop=True)

# ============================================================
# FUNGSI KONVERSI FORMAT ANGKA (US → ID)
# ============================================================
def convert_to_indonesian_format(value):
    """
    Konversi format angka dari 222,432.00 menjadi 222.432,00
    """
    if pd.isna(value) or value == '':
        return ''
    
    value_str = str(value).strip()
    
    # Hapus whitespace dan karakter non-numerik kecuali . dan ,
    value_str = re.sub(r'[^\d,.-]', '', value_str)
    
    if not value_str or value_str == '-':
        return ''
    
    try:
        # Deteksi format: jika ada koma DAN titik, asumsi format US
        if ',' in value_str and '.' in value_str:
            # Format US: 222,432.00 → hapus koma, lalu konversi
            value_str = value_str.replace(',', '')
            number = float(value_str)
        elif ',' in value_str:
            # Hanya ada koma, bisa jadi format ID (222.432,00) atau US (222,432)
            # Cek posisi koma: jika di 3 digit dari belakang = desimal ID
            if value_str.index(',') == len(value_str) - 3:
                # Format ID: 222.432,00
                value_str = value_str.replace('.', '').replace(',', '.')
                number = float(value_str)
            else:
                # Format US: 222,432
                value_str = value_str.replace(',', '')
                number = float(value_str)
        else:
            # Hanya titik atau tanpa separator
            number = float(value_str)
        
        # Format ke Indonesia: 222.432,00
        # Pisahkan bagian integer dan desimal
        formatted = f"{number:,.2f}"  # Dapatkan format US dulu
        
        # Tukar . dan ,
        # 222,432.00 → 222.432,00
        formatted = formatted.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
        
        return formatted
        
    except (ValueError, AttributeError):
        return value_str

# Terapkan konversi pada kolom Debit, Kredit, dan Saldo
print("🔄 Converting number format to Indonesian...\n")
df['Debit'] = df['Debit'].apply(convert_to_indonesian_format)
df['Kredit'] = df['Kredit'].apply(convert_to_indonesian_format)
df['Saldo'] = df['Saldo'].apply(convert_to_indonesian_format)

print(f"📊 Final data: {len(df):,} rows\n")

# Save to Excel
output = 'rekening_koran.xlsx'
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    pd.DataFrame([
        ['Periode', info.get('Periode', '')],
        ['Nama', info.get('Nama', '')],
        ['No Rek', info.get('No_Rek', '')],
        ['Total', f"{len(df):,}"],
    ]).to_excel(writer, sheet_name='Info', index=False, header=False)
    
    df.to_excel(writer, sheet_name='Transaksi', index=False)

print("✅ Excel saved\n")

# Preview
print("="*70)
print("PREVIEW")
print("="*70)
print(df.head(10))
print("\n...")
print(df.tail(10))
print(f"\nTotal: {len(df):,} rows")
print("="*70)

# Download
files.download(output)
df.to_csv('rekening_koran.csv', index=False, encoding='utf-8-sig')
files.download('rekening_koran.csv')

print(f"\n🎉 Done! {len(df):,} transaksi")