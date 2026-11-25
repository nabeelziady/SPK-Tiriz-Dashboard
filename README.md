# Dashboard SPK Tiriz (Python + Streamlit)

## 1. Isi folder
- `app.py` : kode utama dashboard.
- `requirements.txt` : daftar library Python.
- `Database_ModelBase_SPK_Tiriz_Revisi_Dosen_v3.xlsx` : sumber data Excel simulasi.

## 2. Persiapan di VS Code
1. Instal Python 3.10 atau lebih baru.
2. Buka folder proyek ini di VS Code.
3. Buka Terminal > New Terminal.
4. Buat virtual environment:

   Windows:
   ```powershell
   py -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

   macOS/Linux:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   ```

5. Instal library:
   ```bash
   pip install -r requirements.txt
   ```

## 3. Menjalankan website dashboard
Di terminal VS Code:
```bash
streamlit run app.py
```
Browser akan terbuka otomatis. Jika tidak, buka alamat lokal yang muncul di terminal, biasanya `http://localhost:8501`.

## 4. Menggunakan file Excel lain
- Secara default, aplikasi membaca file Excel v3 yang berada satu folder dengan `app.py`.
- Gunakan menu **Unggah file Excel database** di sidebar jika ingin membaca file Excel lain.
- File harus memiliki sheet:
  - `PERGERAKAN_STOK`
  - `TRANSAKSI_PEMAKAIAN`
  - `TRANSAKSI_PEMBELIAN`
  - `MASTER_PARAMETER`

## 5. Rumus Model Base
- `K = D × P`
- `Q = max(0, K − S)`
- `X = max(0, C − S)`
- `Q rekomendasi = min(Q, X)`
- `B = Q rekomendasi × H`

Harga mengubah estimasi biaya, bukan jumlah liter rekomendasi. Batas anggaran hanya menghasilkan peringatan; aplikasi tidak otomatis mengurangi jumlah pembelian.

## 6. Catatan
Data dalam workbook adalah data dummy/simulasi, bukan data aktual UMKM Tiriz. Kapasitas penyimpanan dan batas anggaran merupakan asumsi yang dapat diubah melalui kalkulator.
