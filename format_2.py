# ============================================================
# KONVERSI REKENING KORAN BSI - FINAL FIXED
# Menangani semua kasus description dengan benar
# ============================================================

!pip install pdfplumber pandas openpyxl -q
print("✅ Libraries installed!\n")

from google.colab import files
print("📤 Upload PDF rekening koran BSI:")
uploaded = files.upload()
pdf_file = list(uploaded.keys())[0]
print(f"✅ {pdf_file}\n")

import pdfplumber
import pandas as pd
import re

print("🔄 Extracting data...")

all_rows = []
info = {}

# ============================================================
# FUNGSI PEMBERSIH KARAKTER
# ============================================================
def clean_text(text):
    """Bersihkan karakter aneh dari text"""
    if pd.isna(text) or text == '':
        return ''
    
    text = str(text)
    text = re.sub(r'[^\x20-\x7E\n]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

# ============================================================
# FUNGSI KONVERSI FORMAT ANGKA (US → ID)
# ============================================================
def convert_to_indonesian_format(value):
    """Konversi format angka dari 222,432.00 menjadi 222.432,00"""
    if pd.isna(value) or value == '':
        return ''
    
    value_str = str(value).strip()
    value_str = re.sub(r'[^\d,.-]', '', value_str)
    
    if not value_str or value_str == '-':
        return ''
    
    try:
        if ',' in value_str and '.' in value_str:
            value_str = value_str.replace(',', '')
            number = float(value_str)
        elif ',' in value_str:
            if value_str.index(',') == len(value_str) - 3:
                value_str = value_str.replace('.', '').replace(',', '.')
                number = float(value_str)
            else:
                value_str = value_str.replace(',', '')
                number = float(value_str)
        else:
            number = float(value_str)
        
        formatted = f"{number:,.2f}"
        formatted = formatted.replace(',', 'TEMP').replace('.', ',').replace('TEMP', '.')
        return formatted
        
    except (ValueError, AttributeError):
        return value_str

# ============================================================
# FUNGSI UNTUK CEK APAKAH STRING ADALAH ANGKA
# ============================================================
def is_number(s):
    """Cek apakah string adalah format angka"""
    s = str(s).strip()
    s_clean = s.replace(',', '').replace('.', '').replace('-', '')
    return s_clean.isdigit() and len(s_clean) > 0

# ============================================================
# FUNGSI PARSING BARIS TRANSAKSI - IMPROVED
# ============================================================
def parse_transaction_line(line):
    """
    Parse baris transaksi dengan validasi lebih ketat
    Format: Date Time | FT Number | Description ... | IDR | Amount | DB/CR | Balance
    """
    
    line = line.strip()
    
    # Pattern untuk tanggal dan waktu di awal
    date_pattern = r'^(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})'
    date_match = re.match(date_pattern, line)
    
    if not date_match:
        return None
    
    date = date_match.group(1)
    rest = line[len(date):].strip()
    
    # Split by whitespace
    parts = rest.split()
    
    if len(parts) < 5:
        return None
    
    # Validasi: harus ada IDR, DB atau CR, dan minimal 2 angka (amount & balance)
    if 'IDR' not in [p.upper() for p in parts]:
        return None
    
    if 'DB' not in [p.upper() for p in parts] and 'CR' not in [p.upper() for p in parts]:
        return None
    
    # PARSING DARI BELAKANG untuk memastikan akurasi
    # 1. Balance (angka terakhir)
    balance = parts[-1]
    if not is_number(balance):
        return None
    
    # 2. DB atau CR (kata sebelum balance)
    db_cr = parts[-2].upper()
    if db_cr not in ['DB', 'CR']:
        return None
    
    # 3. Amount (angka sebelum DB/CR)
    amount = parts[-3]
    if not is_number(amount):
        return None
    
    # 4. Currency harus IDR (sebelum amount)
    if parts[-4].upper() != 'IDR':
        return None
    
    # 5. FT Number (elemen pertama setelah date)
    ft_number = parts[0]
    
    # 6. Description = SEMUA yang ada antara FT Number dan IDR
    # Index IDR adalah -4 dari belakang
    idr_position = len(parts) - 4
    
    # Description parts dimulai dari index 1 (setelah FT Number) sampai sebelum IDR
    description_parts = parts[1:idr_position]
    description = ' '.join(description_parts)
    
    # Bersihkan description dari karakter backslash di FT Number jika ada
    # (kadang FT Number punya \BNK di belakangnya yang ikut ke description)
    
    # Tentukan Debit atau Credit
    db_val = 'DB' if db_cr == 'DB' else ''
    cr_val = 'CR' if db_cr == 'CR' else ''
    
    return [date, ft_number, description, 'IDR', amount, db_val, cr_val, balance]

with pdfplumber.open(pdf_file) as pdf:
    total_pages = len(pdf.pages)
    print(f"📄 Total pages: {total_pages}\n")
    
    # Get header info
    first_text = pdf.pages[0].extract_text()
    
    for line in first_text.split('\n'):
        if 'Account' in line and ':' in line and 'Statement' not in line:
            info['Account'] = line.split(':', 1)[-1].strip()
        elif line.startswith('Date') and ':' in line:
            info['Periode'] = line.split(':', 1)[-1].strip()
        elif 'Opening Balance' in line and ':' in line:
            info['Opening_Balance'] = line.split(':', 1)[-1].strip()
        elif 'Closing Balance' in line and ':' in line:
            info['Closing_Balance'] = line.split(':', 1)[-1].strip()
        elif 'Total Debit Amount' in line and ':' in line:
            info['Total_Debit'] = line.split(':', 1)[-1].strip()
        elif 'Total Credit Amount' in line and ':' in line:
            info['Total_Credit'] = line.split(':', 1)[-1].strip()
        elif 'Branch' in line and ':' in line:
            info['Branch'] = line.split(':', 1)[-1].strip()
    
    # Process each page
    for page_num, page in enumerate(pdf.pages, 1):
        print(f"Processing page {page_num}/{total_pages}...", end='\r')
        
        text = page.extract_text()
        
        if not text:
            continue
            
        lines = text.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Skip kosong
            if not line:
                continue
            
            # Skip header lines
            if 'Date' in line and 'FT Number' in line and 'Description' in line:
                continue
            if 'Account Statement' in line:
                continue
            if line.startswith('Page ') and '/' in line:
                continue
            if 'PT ASIA BARU BERKAH MAKASSAR' in line:
                continue
            if line.startswith('Date') and line.endswith('Balance'):
                continue
            
            # Parse baris transaksi
            parsed = parse_transaction_line(line)
            
            if parsed:
                all_rows.append(parsed)
                
                # Debug: print beberapa baris pertama untuk verifikasi
                if len(all_rows) <= 3:
                    print(f"\nDebug row {len(all_rows)}: {parsed[2][:80]}...")  # Print description

print(f"\n✅ Extracted: {len(all_rows):,} rows total\n")

# Create DataFrame
df = pd.DataFrame(all_rows, columns=[
    'Date', 'FT Number', 'Description', 'Currency', 
    'Amount', 'DB', 'CR', 'Balance'
])

# Clean
df = df.replace('', pd.NA)
df = df.dropna(how='all')
df = df.drop_duplicates()
df = df.reset_index(drop=True)

print(f"🔄 Cleaning and formatting data...\n")

# Bersihkan Description
df['Description'] = df['Description'].apply(clean_text)

# Konversi Amount dan Balance
df['Amount'] = df['Amount'].apply(convert_to_indonesian_format)
df['Balance'] = df['Balance'].apply(convert_to_indonesian_format)

# Isi kolom Debit/Credit otomatis
df['Debit'] = ''
df['Credit'] = ''

for idx, row in df.iterrows():
    amount_val = row['Amount']
    db_marker = str(row['DB']).strip().upper()
    cr_marker = str(row['CR']).strip().upper()
    
    if db_marker == 'DB':
        df.at[idx, 'Debit'] = amount_val
        df.at[idx, 'Credit'] = ''
    elif cr_marker == 'CR':
        df.at[idx, 'Debit'] = ''
        df.at[idx, 'Credit'] = amount_val
    else:
        df.at[idx, 'Credit'] = amount_val
        df.at[idx, 'Debit'] = ''

# Hapus kolom Amount, DB, CR
df = df.drop(columns=['Amount', 'DB', 'CR'])

# Reorder kolom
df = df[['Date', 'FT Number', 'Description', 'Currency', 'Debit', 'Credit', 'Balance']]

# Konversi info
for key in ['Opening_Balance', 'Closing_Balance', 'Total_Debit', 'Total_Credit']:
    if key in info:
        info[key] = convert_to_indonesian_format(info[key])

print(f"📊 Final data: {len(df):,} rows\n")

# Save to Excel
output = 'rekening_koran_bsi.xlsx'
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    info_data = [
        ['Account', info.get('Account', '')],
        ['Periode', info.get('Periode', '')],
        ['Branch', info.get('Branch', '')],
        ['Opening Balance', info.get('Opening_Balance', '')],
        ['Closing Balance', info.get('Closing_Balance', '')],
        ['Total Debit', info.get('Total_Debit', '')],
        ['Total Credit', info.get('Total_Credit', '')],
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
print("PREVIEW TRANSAKSI (15 Pertama) - Cek Description Lengkap")
print("="*140)
if len(df) > 0:
    # Tampilkan dengan full description untuk verifikasi
    pd.set_option('display.max_colwidth', None)
    pd.set_option('display.width', None)
    print(df.head(15)[['Date', 'Description', 'Debit', 'Credit']].to_string(index=False))
    print("\n" + "="*140)
else:
    print("❌ Tidak ada data transaksi yang berhasil di-extract!")

print(f"\n{'='*140}")
print(f"Total Transaksi: {len(df):,} rows")
print("="*140)

# Download
if len(df) > 0:
    files.download(output)
    df.to_csv('rekening_koran_bsi.csv', index=False, encoding='utf-8-sig')
    files.download('rekening_koran_bsi.csv')
    print(f"\n🎉 Done! {len(df):,} transaksi berhasil dikonversi")
    print("📥 File downloaded: rekening_koran_bsi.xlsx & rekening_koran_bsi.csv")
else:
    print("\n⚠️ Tidak ada data untuk di-download.")