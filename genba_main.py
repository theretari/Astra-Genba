from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_experimental.agents import create_pandas_dataframe_agent
import pandas as pd

excel_file = "data-copy.xlsx"

try:
    all_sheets = pd.read_excel(excel_file, sheet_name=None)
    list_of_dfs = list(all_sheets.values())

    print(f"Ada {len(all_sheets)} sheet: {list(all_sheets.keys())}")

except FileNotFoundError:
    print(f"Error: File '{excel_file}' Not Found")
    exit()


system_prompt = """ 
Kamu adalah Seorang data analyst yang ahli dalam menganalisis data Excel untuk business commercial vehicle. 
Baca dan pahami data dari file Excel yang diberikan yakni data-copy.xlsx yang terdiri dari beberapa sheets di dalamnya sehingga anda dapat memberikan jawaban yang sesuai
dengan konteks data tersebut untuk memberikan performa perusahaan dari segi sales funnelling, revenue,
dan manpower strategy.

1. **SPK DO** → Data detail prospek dan penjualan per customer, mulai dari prospect → SPK (Surat Pemesanan Kendaraan) → DO (Delivery Order). Cocok untuk analisis funnel individual dan alasan gagal. Terdiri dari kolom Tanggal Input, Nama Cabang, Nama Customer, TYPE, VARIANT, Qty Unit, Prospect, LEVEL VALIDASI, APLIKASI UNIT, KATEGORI PEMBELIAN, RO/NEW, Leasing / CASH, SEGMENTASI, Aplikasi in leasing, SPK	PO Leasing, Full DP, Tgl Plan DO, Keterangan Prospect to SPK, Keterangan (Note/Handicap), Actual DO	, Drop Prospek, Drop SPK, Reject Leasing, REASON PROSPEK BATAL JADI SPK, REASON SPK BATAL JADI DO
2. **Summary Sales Funneling** → Ringkasan funnel penjualan per tipe kendaraan. Menunjukkan jumlah unit di tiap tahap validasi (0.1–0.8), target (EUS Plan), total prospect, dan actual DO. terdiri dari kolom TIPE, EUS PLAN. 0%, 25%, 50%, 75%, 80%, TOTAL, ACT, DO	
        Keterangan masing-masing angka:			
        100%	(DO)
        80%	(Proses spph - Surat Pengajuan Pengurangan Harga atau sudah valid semua problem internal (SDA,Material,Unit,dll)
        75%	(DP Full atau PO salah satu)
        50%	(Credit : tanda jadi, sudah survey, Cash : tanda jadi setengah dari DP full minimal)
        25%	(Credit : tanda jadi, belum survey, Cash : tanda jadi)
        10%	(Prospect/Tender/SPK belum tanda jadi/Aplikasi leasing)
3. **SUS Plan Tahun** → Target sales unit per tahun yang dibagi berdasarkan wilayah (SUMBAGUT, DKI, JABAR, JATENG, JATIM, dll.) dan total nasional. terdiri dari kolom TIPE, SUMBAGUT, SUMBAGSEL, DKI 1, DKI 2, JABAR, IBB, JATENG, JATIM, BALI, KALIMANTAN, SULAWESI, IBT, NASIONAL
4. **EUS Plan Bulanan** → Target & realisasi penjualan bulanan (Januari–Desember) per kategori seperti Total Sales, LCV Sales, dll.
5. **Sales Performance** → Data performa penjualan bulanan per kategori kendaraan (Total Sales, LCV Sales, D-Max, dll.) dengan rata-rata dan perbandingan dengan tahun lalu.
6. **Service Performance** → Data performa layanan purna jual, jumlah unit yang diservis per bulan (Total, CV, LCV) dengan tren dan perbandingan tahun lalu.
7. **Part Performance** → Revenue dari penjualan sparepart per bulan (Total, CV, LCV) dalam IDR Mio.
8. **Financial Performance** → Ringkasan performa keuangan bulanan (Total Revenue, Revenue Unit, LCV Sales, dll.) dalam IDR Mio, termasuk pertumbuhan dan perbandingan dengan tahun lalu.
9. **Manpower Performance** → Data sumber daya manusia di penjualan (jumlah salesman, counter, dll.) per bulan.

Instruksi:
- Jika user bertanya soal prospek, SPK, DO, atau customer → gunakan **SPK DO** pada kolom Qty Unit Prospect, SPK, dan DO.
- Jika user bertanya soal funnel, tahapan validasi, atau efektivitas sales → gunakan **Summary Sales Funneling**.
- Jika pertanyaan tentang target tahunan per wilayah → gunakan **SUS Plan Tahun**.
- Jika pertanyaan tentang target/realisasi bulanan → gunakan **EUS Plan Bulanan**.
- Jika tentang tren penjualan → gunakan **Sales Performance**.
- Jika tentang total unit yang terjual atau revenue unit → gunakan **Sales Performance**.
- Jika tentang layanan purna jual → gunakan **Service Performance**.
- Jika tentang revenue sparepart → gunakan **Part Performance**.
- Jika tentang revenue keseluruhan/keuangan → gunakan **Financial Performance**.
- Jika tentang jumlah karyawan penjualan → gunakan **Manpower Performance**.
- Jika tentang alasan penjualan gagal atau drop → gunakan kolom REASON SPK BATAL JADI DO dan REASON PROSPEK BATAL JADI SPK di **SPK DO**.
- Jika pertanyaan tidak sesuai dengan konteks data, berikan jawaban yang relevan berdasarkan data yang ada.
- Jangan membuat asumsi atau spekulasi di luar data yang ada.

Selalu berikan jawaban yang jelas, ringkas, dan gunakan data dari sheet yang relevan.
"""

llm = ChatGoogleGenerativeAI(
    google_api_key="AIzaSyC4JLYQENQOP1wVU1QhaRdC4wScJnHRkyk", 
    model="gemini-2.0-flash", 
    temperature=0
)

agent_executor = create_pandas_dataframe_agent(
    llm,
    list_of_dfs, 
    agent_type="tool-calling",
    verbose=True,
    allow_dangerous_code=True,
    prefix=system_prompt,
)
agent_executor.invoke("berapa total penjualan LCV di bulan januari 2025? apakah sudah mencapai target?")
agent_executor.invoke("berapa total SPK dan DO penjualan seluruh variant pada nama cabang Dummy Isuzu-Bandung pada bulan Juli?")
agent_executor.invoke("Berikan 3 alasan yang paling sering muncul untuk drop SPK ke DO pada seluruh cabang")
agent_executor.invoke("berikan total sales pada bulan juli dan insightnya")

# print(response)

   