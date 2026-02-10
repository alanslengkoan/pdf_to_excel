# ============================================================
# KONVERSI REKENING KORAN BCA
# Format: Tanggal | Keterangan | CBG | Mutasi | Saldo
# ============================================================

!pip install pdfplumber pandas openpyxl -q
print("✅ Libraries installed!\n")

from google.colab import files
print("📤 Upload PDF rekening koran BCA:")
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
    """Konversi format angka dari 196,000.00 menjadi 196.000,00"""
    if pd.isna(value) or value == '' or value == '0.00':
        return ''
    
    value_str = str(value).strip()
    value_str = re.sub(r'[^\d,.-]', '', value_str)
    
    if not value_str or value_str == '-' or value_str == '0.00':
        return ''
    
    try:
        # Format BCA: 196,000.00 (koma ribuan, titik desimal)
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
        
        # Format ke Indonesia: 196.000,00
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
    
    # Get header info dari halaman pertama
    first_text = pdf.pages[0].extract_text()
    lines = first_text.split('\n')
    
    for i, line in enumerate(lines[:40]):
        # Extract nama (biasanya baris pertama dalam box)
        if i > 5 and i < 15 and len(line.strip()) > 10:
            if not any(x in line.upper() for x in ['REKENING', 'HALAMAN', 'PERIODE', 'NO.', 'CATATAN']):
                if 'Nama' not in info:
                    info['Nama'] = line.strip()
        
        # No Rekening
        if 'NO. REKENING' in line.upper() or 'NO.REKENING' in line.upper():
            match = re.search(r':?\s*(\d+)', line)
            if match:
                info['No_Rekening'] = match.group(1)
        
        # Periode
        if 'PERIODE' in line.upper():
            match = re.search(r':?\s*([A-Z]+\s+\d{4})', line, re.IGNORECASE)
            if match:
                info['Periode'] = match.group(1)
        
        # Mata Uang
        if 'MATA UANG' in line.upper():
            if 'IDR' in line:
                info['Mata_Uang'] = 'IDR'
    
    print("🔍 Extracting transactions...\n")
    
    # Process each page
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"   Processing page {page_num}/{total_pages}...", end='\r')
        
        # Extract text
        text = page.extract_text()
        
        if not text:
            continue
        
        lines = text.split('\n')
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            # Skip kosong
            if not line:
                i += 1
                continue
            
            # Skip header
            if 'TANGGAL' in line.upper() and 'KETERANGAN' in line.upper():
                i += 1
                continue
            if 'REKENING GIRO' in line:
                i += 1
                continue
            if 'Bersambung' in line:
                i += 1
                continue
            
            # Cek apakah baris dimulai dengan tanggal (format: DD/MM)
            date_match = re.match(r'^(\d{2}/\d{2})\s+(.+)', line)
            
            if date_match:
                tanggal = date_match.group(1)
                rest = date_match.group(2).strip()
                
                # Kumpulkan baris-baris berikutnya yang merupakan bagian dari transaksi ini
                # (baris yang tidak dimulai dengan tanggal atau SALDO AWAL)
                continuation_lines = []
                j = i + 1
                
                while j < len(lines):
                    next_line = lines[j].strip()
                    
                    # Stop jika ketemu tanggal baru atau SALDO AWAL
                    if re.match(r'^\d{2}/\d{2}\s+', next_line):
                        break
                    if next_line.startswith('SALDO AWAL'):
                        break
                    if not next_line:
                        j += 1
                        continue
                    
                    # Tambahkan ke continuation
                    continuation_lines.append(next_line)
                    j += 1
                
                # Gabungkan semua baris
                full_text = rest + ' ' + ' '.join(continuation_lines)
                full_text = re.sub(r'\s+', ' ', full_text).strip()
                
                # Parse data
                # Format: Keterangan ... CBG ... Mutasi ... [Saldo]
                # Cari semua angka dalam format amount
                amounts = re.findall(r'\d{1,3}(?:,\d{3})*(?:\.\d{2})', full_text)
                
                if len(amounts) == 0:
                    i = j
                    continue
                
                # Identifikasi Saldo dan Mutasi
                # Saldo biasanya angka terbesar dan ada di akhir
                # Mutasi adalah angka sebelum saldo
                
                saldo = ''
                mutasi = ''
                
                if len(amounts) >= 2:
                    mutasi = amounts[-2]
                    saldo = amounts[-1]
                elif len(amounts) == 1:
                    # Jika hanya 1 angka, bisa jadi saldo saja atau mutasi saja
                    # Check apakah ada kata kunci
                    if 'SALDO AWAL' in full_text.upper():
                        saldo = amounts[0]
                    else:
                        mutasi = amounts[0]
                
                # Extract Keterangan dan CBG
                # Hapus amounts dari text
                text_only = full_text
                for amt in amounts:
                    text_only = text_only.replace(amt, '', 1)
                
                text_only = re.sub(r'\s+', ' ', text_only).strip()
                
                # CBG biasanya kode singkat (DR 564, dll)
                # atau kode transaksi
                cbg = ''
                keterangan_parts = []
                
                parts = text_only.split()
                
                # Cari pattern CBG (biasanya 2-3 kata di tengah atau akhir)
                for idx, part in enumerate(parts):
                    # DR, CR diikuti angka
                    if part in ['DR', 'CR'] and idx + 1 < len(parts):
                        cbg = f"{part} {parts[idx + 1]}"
                        break
                
                # Keterangan adalah sisanya
                if cbg:
                    keterangan = text_only.replace(cbg, '', 1).strip()
                else:
                    # Ambil beberapa kata terakhir sebagai CBG jika ada
                    if len(parts) > 3:
                        # Coba ambil 2 kata terakhir sebagai CBG
                        potential_cbg = ' '.join(parts[-2:])
                        # Cek apakah terlihat seperti kode
                        if len(potential_cbg) < 20 and (potential_cbg.isupper() or re.match(r'^[\w\s-]+$', potential_cbg)):
                            cbg = potential_cbg
                            keterangan = ' '.join(parts[:-2])
                        else:
                            keterangan = text_only
                    else:
                        keterangan = text_only
                
                keterangan = clean_text(keterangan)
                cbg = clean_text(cbg)
                
                all_rows.append([tanggal, keterangan, cbg, mutasi, saldo])
                
                i = j
            else:
                i += 1

print(f"\n\n✅ Extracted: {len(all_rows):,} rows\n")

if len(all_rows) == 0:
    print("❌ Tidak ada data yang terekstrak!")
    print("\nDEBUG: Sample text untuk analisa:")
    with pdfplumber.open(pdf_file) as pdf:
        text = pdf.pages[0].extract_text()
        lines = text.split('\n')
        print("\nBaris yang mengandung tanggal:")
        for i, line in enumerate(lines):
            if re.search(r'\d{2}/\d{2}', line):
                print(f"   {i}: {line}")
else:
    # Create DataFrame
    df = pd.DataFrame(all_rows, columns=[
        'Tanggal', 'Keterangan', 'CBG', 'Mutasi', 'Saldo'
    ])
    
    df = df.replace('', pd.NA)
    df = df.dropna(how='all')
    df = df.drop_duplicates()
    df = df.reset_index(drop=True)
    
    print(f"🔄 Cleaning and formatting data...\n")
    
    df['Keterangan'] = df['Keterangan'].fillna('').apply(clean_text)
    df['CBG'] = df['CBG'].fillna('').apply(clean_text)
    
    df['Mutasi'] = df['Mutasi'].fillna('').apply(convert_to_indonesian_format)
    df['Saldo'] = df['Saldo'].fillna('').apply(convert_to_indonesian_format)
    
    # Hitung total mutasi
    total_mutasi = df['Mutasi'].apply(lambda x: float(x.replace('.', '').replace(',', '.')) if x else 0).sum()
    
    info['Total_Mutasi'] = f"{total_mutasi:,.2f}".replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
    
    print(f"📊 Final data: {len(df):,} rows\n")
    
    # Save to Excel
    output = 'rekening_koran_bca.xlsx'
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        info_data = [
            ['Nama', info.get('Nama', '')],
            ['No. Rekening', info.get('No_Rekening', '')],
            ['Periode', info.get('Periode', '')],
            ['Mata Uang', info.get('Mata_Uang', '')],
            ['Total Mutasi', info.get('Total_Mutasi', '')],
            ['Total Transaksi', f"{len(df):,}"],
        ]
        pd.DataFrame(info_data).to_excel(writer, sheet_name='Info', index=False, header=False)
        df.to_excel(writer, sheet_name='Transaksi', index=False)
    
    print("✅ Excel saved\n")
    
    # Preview
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
    print(f"   Total Transaksi: {len(df):,}")
    
    print(f"\n{'='*140}")
    
    files.download(output)
    df.to_csv('rekening_koran_bca.csv', index=False, encoding='utf-8-sig')
    files.download('rekening_koran_bca.csv')
    print(f"\n🎉 Done! {len(df):,} transaksi berhasil dikonversi")
    print("📥 File: rekening_koran_bca.xlsx & rekening_koran_bca.csv")