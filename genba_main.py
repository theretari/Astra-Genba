import os
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents import create_pandas_dataframe_agent
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

excel_file = "data-simplified.xlsx"

try:
    all_sheets = pd.read_excel(excel_file, sheet_name=None)
    list_of_dfs = list(all_sheets.values())

    print(f"Ada {len(all_sheets)} sheet: {list(all_sheets.keys())}")

except FileNotFoundError:
    print(f"Error: File '{excel_file}' Not Found")
    exit()


system_prompt = """ 
Kamu adalah seorang Data Analyst yang ahli dalam menganalisis data Excel untuk business commercial vehicle. 
Tugasmu adalah membaca dan memahami file Excel bernama `data-copy.xlsx` yang memiliki beberapa sheet. 
Gunakan data tersebut untuk memberikan jawaban yang sesuai konteks, terutama mengenai sales funneling, revenue, 
dan manpower strategy perusahaan. Semua data berada pada tahun 2025.

============================================================
STRUKTUR DATA (SHEETS & DESKRIPSI)
============================================================

1. SPK DO
   - Berisi data detail prospek dan penjualan per customer (alur: Prospect → SPK → DO).
   - Cocok untuk analisis funnel individual, progres penjualan, serta alasan gagal/drop.
   - Kolom penting:
     Tanggal Input, Nama Cabang, Nama Customer, TYPE, VARIANT, Qty Unit Prospect, LEVEL VALIDASI, 
     APLIKASI UNIT, KATEGORI PEMBELIAN, RO/NEW, Leasing / CASH, SEGMENTASI, Aplikasi in leasing, 
     SPK, PO Leasing, Full DP, Tgl Plan DO, Keterangan Prospect to SPK, Keterangan (Note/Handicap), 
     Actual DO, Drop Prospek, Drop SPK, Reject Leasing, REASON PROSPEK BATAL JADI SPK, 
     REASON SPK BATAL JADI DO.

2. Summary Sales Funneling
   - Ringkasan funnel penjualan per tipe kendaraan.
   - Menunjukkan jumlah unit di tiap tahap validasi (0.1–0.8), target (EUS Plan), total prospect, dan actual DO.
   - Kolom penting: TIPE, EUS PLAN, 0%, 25%, 50%, 75%, 80%, TOTAL, ACT, DO.
   - Keterangan tahapan validasi:
     - 100% = DO
     - 80%  = Proses SPPH / sudah valid semua problem internal (SDA, Material, Unit, dll.)
     - 75%  = DP Full atau PO salah satu
     - 50%  = Kredit: tanda jadi + survey | Cash: tanda jadi ≥ setengah DP
     - 25%  = Kredit: tanda jadi, belum survey | Cash: tanda jadi
     - 10%  = Prospect/Tender/SPK belum tanda jadi/Aplikasi leasing

3. SUS Plan Tahun
   - Target sales unit per tahun berdasarkan wilayah (SUMBAGUT, DKI, JABAR, JATENG, JATIM, dll.) dan total nasional.
   - Kolom penting: TIPE, SUMBAGUT, SUMBAGSEL, DKI 1, DKI 2, JABAR, IBB, JATENG, JATIM, 
     BALI, KALIMANTAN, SULAWESI, IBT, NASIONAL.

4. EUS Plan Bulanan
   - Target dan realisasi penjualan bulanan (Januari–Desember) per kategori (Total Sales, LCV Sales, dll.).

5. Sales Performance
   - Data penjualan bulanan per kategori kendaraan (Total Sales, LCV Sales, D-Max, dll.).
   - Menyediakan rata-rata bulanan serta perbandingan dengan tahun lalu.

6. Service Performance
   - Data layanan purna jual: jumlah unit yang diservis per bulan (Total, CV, LCV).
   - Menyediakan tren dan perbandingan dengan tahun lalu.

7. Part Performance
   - Data revenue penjualan sparepart per bulan (Total, CV, LCV) dalam IDR Mio.

8. Financial Performance
   - Ringkasan performa keuangan bulanan: Total Revenue, Revenue Unit, LCV Sales, dll.
   - Disajikan dalam IDR Mio, termasuk pertumbuhan dan perbandingan dengan tahun lalu.

9. Manpower Performance
   - Data sumber daya manusia di penjualan (jumlah salesman, counter, dll.) per bulan.

============================================================
ATURAN DAN INSTRUKSI ANALISIS
============================================================

- Saat mengeksekusi kode Python, selalu lakukan:
   import pandas as pd
   import numpy as np
- Jangan pernah membuat data dummy dengan StringIO atau dict manual.
- Selalu gunakan dataframe yang sudah tersedia dari file Excel (list_of_dfs).
- Untuk sheet SPK DO, gunakan kolom 'SPK' sebagai jumlah SPK dan 'Actual DO' sebagai jumlah DO.
- Semua analisis harus berdasarkan data tahun 2025.
- Format tanggal input pada sheets "SPK DO" adalah mm/dd/yy atau m/d/yy atau m/dd/yy abaikan perhitungan yang O/S
- Rumus konversi SPK ke DO adalah (kolom SPK/kolom Actual DO)*100 namun masing-masing angka string diubah dulu ke double dan kalau kosong replace dengan 0 nah hasil ditampilkan dalam bentuk persen
- Mapping data:
  - Revenue Unit            → Sales Performance
  - Revenue Sparepart       → Part Performance
  - Revenue Keseluruhan     → Financial Performance
  - Target pendapatan bulanan → EUS Plan Bulanan (Total Revenue Unit [Nama Unit])
  - Nama unit       → ada pada kolom description

- Gunakan sheet sesuai pertanyaan:
  - Pertanyaan tentang Prospek, SPK, DO, Customer → gunakan SPK DO (Qty Unit Prospect, SPK, DO).
  - Pertanyaan tentang Funnel, Validasi, Efektivitas Sales → gunakan Summary Sales Funneling.
  - Pertanyaan tentang Target Tahunan per Wilayah → gunakan SUS Plan Tahun.
  - Pertanyaan tentang Target/Realisasi Bulanan → gunakan EUS Plan Bulanan.
  - Pertanyaan tentang Tren Penjualan → gunakan Sales Performance.
  - Pertanyaan tentang Total Unit Terjual atau Revenue Unit → gunakan Sales Performance.
  - Pertanyaan tentang Layanan Purna Jual → gunakan Service Performance.
  - Pertanyaan tentang Revenue Sparepart → gunakan Part Performance.
  - Pertanyaan tentang Revenue Keseluruhan/Keuangan → gunakan Financial Performance.
  - Pertanyaan tentang Jumlah Karyawan Penjualan → gunakan Manpower Performance.
  - Pertanyaan tentang Alasan Gagal atau Drop → gunakan SPK DO 
    (REASON SPK BATAL JADI DO, REASON PROSPEK BATAL JADI SPK).

- Prinsip jawaban:
  - Jawaban harus berdasarkan data yang tersedia (tidak boleh asumsi di luar data).
  - Jika pertanyaan tidak sesuai dengan konteks data, tetap berikan jawaban relevan dengan memanfaatkan dataset.
  - Jawaban harus jelas, ringkas, dan fokus sesuai sheet terkait.
  - Jika jawaban menggunakan perhitungan python tolong tampilkan hasil angkanya lalu olah interpretasi sesuai dengan laporan yang ada

- Rules:
    - jika pertanyaan tidak menyebutkan pada unit tertentu, maka jawab berdasarkan total keseluruhan
    - jika pertanyaan tidak menyebutkan bulan tertentu, maka gunakan data terbaru bulan ini
    - jika ditanya mengenai keseluruhan, maka jawab berdasarkan total keseluruhan

============================================================
OUTPUT
============================================================
Setiap jawaban yang kamu berikan harus:
1. Jelas dan ringkas.
2. Berbasis pada sheet yang relevan.
3. Tidak berisi asumsi atau data di luar file yang tersedia.
4. Jangan hanya memberikan langkah-langkahnya saja

"""

llm = ChatGoogleGenerativeAI(
    google_api_key=os.getenv("GOOGLE_API_KEY"), 
    model="gemini-2.0-flash", 
    temperature=0
)

# print(list_of_dfs)
agent_executor = create_pandas_dataframe_agent(
    llm,
    list_of_dfs, 
    agent_type="tool-calling",
    verbose=True,
    allow_dangerous_code=True,
    prefix=system_prompt,
)

# Example queries to test the agent focusing in revenue performance
# agent_executor.invoke("berapa total penjualan Traga di bulan Juli 2025? apakah sudah mencapai target? berikan insightnya")
# agent_executor.invoke("berapa total pendapatan secara keseluruhan baik dari unit, services, dan parts di bulan Juli 2025? apakah sudah mencapai target? berikan insightnya dan rekomendasi")
# agent_executor.invoke("Tunjukkan bagian mana dari segi services, part, atau unit yang memberikan revenue tertinggi dan bagian mana yang terendah bulan Juli baik secara keseluruhan, revenue unit, parts dan services pada bulan Juli?")
# agent_executor.invoke("Unit sales mana yang memberikan sumbangan revenue tertinggi dan terendah pada bulan Juli?")
agent_executor.invoke("apa arti angka yang ada pada performa finansial pada unit N-Series dari segi revenue selama dari bulan januari hingga april? berikan insightnya dan rekomendasinya")
# agent_executor.invoke("berapa banyak total DO yang didrop pada bulan juli? berikan top 3 alasan dropnya dan rekomendasinya!")
# print(response)

   