# run ini di terminal sebelum run code:
# pip install pandas openpyxl
import pandas as pd
import glob
import os
import datetime as dt

# ==============================================================================
# LANGKAH 1: MENEMUKAN DAN MENGGABUNGKAN SEMUA FILE TRANSAKSI
# ==============================================================================
path_ke_dataset = '../dataset/*.xlsx'
file_paths = glob.glob(path_ke_dataset)
if not file_paths:
    print("Tidak ada file Excel (.xlsx) yang ditemukan di folder dataset.")
else:
    print(f"Ditemukan {len(file_paths)} file untuk digabungkan:")
    for path in file_paths:
        print(f"- {os.path.basename(path)}")
    list_of_dataframes = []
    for path in file_paths:
        try:
            df = pd.read_excel(path, skiprows=4)
            list_of_dataframes.append(df)
        except Exception as e:
            print(f"Error saat membaca file {path}: {e}")
    if list_of_dataframes:
        merged_df = pd.concat(list_of_dataframes, ignore_index=True)
        print("\n✅ Semua file berhasil digabungkan.")
        print("Total baris data setelah digabung:", len(merged_df))
        print("\nKolom yang tersedia di data Anda:", merged_df.columns.tolist())
        
        # ==============================================================================
        # LANGKAH 2: PEMBERSIHAN DAN PERSIAPAN DATA UNTUK RFM
        # ==============================================================================
        print("\nMemulai persiapan data untuk RFM...")
        KOLOM_TANGGAL = 'Tgl. Trans.'
        KOLOM_INVOICE = 'No. Trans.'
        KOLOM_PELANGGAN = 'Customer'
        KOLOM_TOTAL_BELANJA = 'Subtotal (Rp)'
        kolom_rfm = [KOLOM_TANGGAL, KOLOM_INVOICE, KOLOM_PELANGGAN, KOLOM_TOTAL_BELANJA]
        if all(col in merged_df.columns for col in kolom_rfm):
            rfm_df = merged_df[kolom_rfm].copy()
            rfm_df.dropna(subset=[KOLOM_PELANGGAN, KOLOM_TANGGAL], inplace=True)
            rfm_df[KOLOM_TANGGAL] = pd.to_datetime(rfm_df[KOLOM_TANGGAL], errors='coerce')
            rfm_df[KOLOM_TOTAL_BELANJA] = pd.to_numeric(rfm_df[KOLOM_TOTAL_BELANJA], errors='coerce').fillna(0)
            rfm_df.dropna(subset=[KOLOM_TANGGAL], inplace=True)
            print("✅ Data berhasil dibersihkan dan disiapkan.")

            # ==============================================================================
            # LANGKAH 3: MENGHITUNG NILAI RECENCY, FREQUENCY, MONETARY
            # ==============================================================================
            print("\nMenghitung nilai R, F, M untuk setiap pelanggan...")
            snapshot_date = rfm_df[KOLOM_TANGGAL].max() + dt.timedelta(days=1)
            print(f"Tanggal acuan (snapshot date): {snapshot_date.date()}")
            tabel_rfm_final = rfm_df.groupby(KOLOM_PELANGGAN).agg({
                KOLOM_TANGGAL: lambda date: (snapshot_date - date.max()).days,
                KOLOM_INVOICE: 'nunique',
                KOLOM_TOTAL_BELANJA: 'sum'
            })
            tabel_rfm_final.rename(columns={
                KOLOM_TANGGAL: 'Recency',
                KOLOM_INVOICE: 'Frequency',
                KOLOM_TOTAL_BELANJA: 'Monetary'
            }, inplace=True)
            print("✅ Perhitungan RFM selesai.")

            # ==============================================================================
            # LANGKAH 4: MENYIMPAN HASIL DAN MENAMPILKAN CONTOH
            # ==============================================================================
            output_filename = 'tabel_rfm.csv'
            tabel_rfm_final.to_csv(output_filename, encoding='utf-8-sig')
            print(f"\n🎉 Berhasil! Tabel RFM telah disimpan sebagai '{output_filename}'.")
            print("\nContoh 5 baris pertama dari tabel RFM:")
            print(tabel_rfm_final.head())
            
            # ==============================================================================
            # LANGKAH 5: EKSPOR TRANSAKSI LENGKAP PELANGGAN MONETARY NOL (FINAL)
            # ==============================================================================
            print("\n" + "="*70)
            print("LANGKAH 5: EKSPOR TRANSAKSI PELANGGAN DENGAN MONETARY NOL")
            print("="*70)

            # 1. Cari nama pelanggan dengan Monetary = 0 dari tabel RFM
            zero_monetary_customers = tabel_rfm_final[tabel_rfm_final['Monetary'] == 0].index.tolist()

            if zero_monetary_customers:
                print(f"\nDitemukan {len(zero_monetary_customers)} pelanggan dengan total belanja nol.")

                # 2. Saring merged_df untuk mendapatkan semua baris transaksi milik pelanggan tersebut
                df_transaksi_zero_monetary = merged_df[merged_df[KOLOM_PELANGGAN].isin(zero_monetary_customers)]
                
                # 3. Simpan hasil saringan ke satu file CSV baru
                output_zero_filename = 'transaksi_pelanggan_monetary_nol.csv'
                df_transaksi_zero_monetary.to_csv(output_zero_filename, index=False, encoding='utf-8-sig')
                
                print(f"✅ Semua baris transaksi dari pelanggan tersebut telah disimpan di file '{output_zero_filename}'.")

            else:
                print("\n✅ Tidak ditemukan pelanggan dengan total belanja nol.")

        else:
            print(f"\n❌ Error: Satu atau lebih kolom yang dibutuhkan untuk RFM tidak ditemukan.")
            print(f"Pastikan file Excel Anda memiliki kolom: {kolom_rfm}")
            