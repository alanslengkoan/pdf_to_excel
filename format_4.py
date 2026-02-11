# Install dependencies di Google Colab
!pip install pdfplumber openpyxl

# Import libraries
import pdfplumber
import pandas as pd
import re
from datetime import datetime
from google.colab import files

def parse_bca_pdf(pdf_path):
    """
    Parse BCA bank statement PDF to extract transactions
    """
    transactions = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            lines = text.split('\n')
            
            for line in lines:
                # Match transaction lines starting with date DD/MM
                if re.match(r'^\d{2}/\d{2}\s+', line):
                    try:
                        # Extract components
                        parts = line.split()
                        date = parts[0]
                        
                        # Find all numbers in the line
                        numbers = re.findall(r'([\d,]+\.\d{2})', line)
                        
                        # Get transaction type
                        trans_type = ''
                        if 'TRSF E-BANKING CR' in line:
                            trans_type = 'Transfer E-Banking (Kredit)'
                        elif 'TRSF E-BANKING DB' in line:
                            trans_type = 'Transfer E-Banking (Debit)'
                        elif 'BI-FAST CR' in line:
                            trans_type = 'BI-FAST (Kredit)'
                        elif 'BI-FAST DB' in line:
                            trans_type = 'BI-FAST (Debit)'
                        elif 'SWITCHING CR' in line:
                            trans_type = 'Switching (Kredit)'
                        elif 'BIAYA ADM' in line:
                            trans_type = 'Biaya Admin'
                        elif 'BUNGA' in line:
                            trans_type = 'Bunga'
                        elif 'PAJAK BUNGA' in line:
                            trans_type = 'Pajak Bunga'
                        else:
                            trans_type = parts[1] if len(parts) > 1 else ''
                        
                        # Extract description
                        # Remove date, transaction type, and numbers
                        desc = line
                        desc = re.sub(r'^\d{2}/\d{2}\s+', '', desc)
                        desc = re.sub(r'(TRSF E-BANKING CR|TRSF E-BANKING DB|BI-FAST CR|BI-FAST DB|SWITCHING CR|BIAYA ADM|BUNGA|PAJAK BUNGA)', '', desc)
                        desc = re.sub(r'\s+[\d,]+\.\d{2}.*$', '', desc)
                        desc = re.sub(r'^\s*\d+\s+', '', desc)  # Remove CBG code
                        desc = desc.strip()
                        
                        # Determine debit/credit and amounts
                        is_debit = ' DB' in line
                        
                        if numbers:
                            balance = float(numbers[-1].replace(',', ''))
                            
                            if len(numbers) >= 2:
                                amount = float(numbers[-2].replace(',', ''))
                            else:
                                amount = 0
                        else:
                            amount = 0
                            balance = 0
                        
                        transaction = {
                            'Tanggal': date + '/2025',
                            'Tipe Transaksi': trans_type,
                            'Keterangan': desc,
                            'Debit': amount if is_debit else 0,
                            'Kredit': amount if not is_debit else 0,
                            'Saldo': balance
                        }
                        
                        transactions.append(transaction)
                    
                    except Exception as e:
                        print(f"Error parsing line: {line[:50]}... - {e}")
                        continue
    
    return transactions

def create_excel_report(transactions, output_path):
    """
    Create Excel file with transactions and summary
    """
    # Create main DataFrame
    df = pd.DataFrame(transactions)
    
    # Calculate summary
    total_kredit = df['Kredit'].sum()
    total_debit = df['Debit'].sum()
    saldo_awal = 645447905.64
    saldo_akhir = df['Saldo'].iloc[-1] if len(df) > 0 else 0
    
    summary_data = {
        'Keterangan': ['Saldo Awal', 'Total Kredit', 'Total Debit', 'Saldo Akhir', 'Jumlah Transaksi'],
        'Nilai': [
            f'Rp {saldo_awal:,.2f}',
            f'Rp {total_kredit:,.2f}',
            f'Rp {total_debit:,.2f}',
            f'Rp {saldo_akhir:,.2f}',
            f'{len(df)} transaksi'
        ]
    }
    df_summary = pd.DataFrame(summary_data)
    
    # Format currency columns
    for col in ['Debit', 'Kredit', 'Saldo']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f'Rp {x:,.2f}' if x > 0 else '')
    
    # Write to Excel
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Ringkasan', index=False)
        df.to_excel(writer, sheet_name='Detail Transaksi', index=False)
        
        # Auto-adjust column width
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return df, df_summary

# MAIN EXECUTION
print("=== BCA PDF to Excel Converter ===\n")

# Upload PDF file
print("Silakan upload file PDF BCA Anda:")
uploaded = files.upload()

# Get the uploaded filename
pdf_filename = list(uploaded.keys())[0]
print(f"\nFile uploaded: {pdf_filename}")

# Parse PDF
print("\nMemproses PDF...")
transactions = parse_bca_pdf(pdf_filename)
print(f"Berhasil mengekstrak {len(transactions)} transaksi")

# Create Excel
excel_filename = 'BCA_Rekening_Koran_Desember_2025.xlsx'
print(f"\nMembuat file Excel: {excel_filename}")
df_trans, df_summary = create_excel_report(transactions, excel_filename)

# Display summary
print("\n=== RINGKASAN ===")
print(df_summary.to_string(index=False))

# Display sample transactions
print("\n=== SAMPLE DATA (10 Transaksi Pertama) ===")
print(df_trans.head(10).to_string(index=False))

# Download the Excel file
print(f"\nMengunduh file Excel...")
files.download(excel_filename)

print("\n✅ Konversi selesai!")