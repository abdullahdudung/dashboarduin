# Dashboard IKU UIN Jakarta — Streamlit

Revisi dashboard E-Semesta untuk kebutuhan pimpinan: ringkasan mahasiswa aktif, profil per fakultas/jenjang, PMB, matriks IKU, dan kesiapan data.

## Menjalankan di Windows 11 / Antigravity IDE

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Buka `http://localhost:8501`.

## Catatan penting

Data saat ini mendukung profil mahasiswa dan agregat PMB per program studi. Beberapa kebutuhan pimpinan belum dapat dihitung secara resmi karena kolom MBKM, kuota/jalur PMB, kelulusan, daerah 3T, top percentile, dan alumni belum tersedia. Dashboard menandainya sebagai **Proxy** atau **Belum tersedia** agar tidak menghasilkan angka yang menyesatkan.
