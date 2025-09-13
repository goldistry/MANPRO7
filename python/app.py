# run ini di terminal sebelum run code:
# pip install pandas openpyxl
import pandas as pd
import glob
import os

# Path relatif dari folder script (python) ke folder dataset
path_ke_dataset = '../dataset/*.xlsx'

# Gunakan path tersebut untuk menemukan semua file Excel
file_paths = glob.glob(path_ke_dataset)

# Cek apakah file ditemukan
if not file_paths:
    print("Tidak ada file Excel (.xlsx) yang ditemukan di folder dataset.")
else:
    print(f"Ditemukan {len(file_paths)} file untuk digabungkan:")
    for path in file_paths:
        print(f"- {os.path.basename(path)}")

    list_of_dataframes = []

    # Loop melalui setiap path file yang ditemukan
    for path in file_paths:
        try:
            # GUNAKAN pd.read_excel UNTUK FILE .xlsx
            df = pd.read_excel(path, skiprows=4)
            list_of_dataframes.append(df)
        except Exception as e:
            print(f"Error saat membaca file {path}: {e}")

    # Lanjutkan hanya jika berhasil membaca setidaknya satu file
    if list_of_dataframes:
        merged_df = pd.concat(list_of_dataframes, ignore_index=True)
        print("\nSemua file berhasil digabungkan.")
        print("Total baris data setelah digabung:", len(merged_df))

        # Definisikan kolom yang ingin kita proses
        kolom_target = ['Judul', 'ISBN', 'ISBN Dtl.']
        
        # Cek apakah semua kolom target ada di DataFrame
        if all(col in merged_df.columns for col in kolom_target):
            # Langsung pilih kolom yang dibutuhkan
            books_df = merged_df[kolom_target]

            # Hapus baris di mana 'Judul' kosong
            books_df = books_df.dropna(subset=['Judul'])
            
            # Hapus baris duplikat
            unique_books_df = books_df.drop_duplicates(subset=kolom_target)

            print(f"\nDitemukan {len(unique_books_df)} data buku yang unik.")

            # --- LANGKAH 4: MEMBUAT FILE CSV BARU UNTUK GENRE ---
            final_df = unique_books_df.copy()
            final_df['Genre'] = ''

            output_filename = 'daftar_buku_unique.csv'
            final_df.to_csv(output_filename, index=False, encoding='utf-8-sig')

            print(f"\nBerhasil! File '{output_filename}' telah dibuat.")
            print("File ini berisi daftar buku unik (Judul, ISBN, ISBN Dtl.).")
        
        else:
            # Pesan error yang relevan
            print(f"\nError: Satu atau lebih kolom target {kolom_target} tidak ditemukan.")
            print("Pastikan semua file Excel Anda memiliki kolom tersebut.")
            print("Kolom yang tersedia:", merged_df.columns.tolist())