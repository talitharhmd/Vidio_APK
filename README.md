# Vidio Plays Dashboard (Streamlit)

Dashboard 2 halaman yang membaca satu tabel: `fact_dashboard.xlsx` (dibuat oleh bagian 10b notebook).

## Isi folder
- `app.py` : kode dashboard
- `requirements.txt` : library yang dibutuhkan
- `.streamlit/config.toml` : tema (merah Vidio)
- `vidio_logo.png`, `vidio_icon.png` : logo
- `fact_dashboard.xlsx` : **kamu yang menambahkan**, ambil dari `vidio_output/dashboard/interaktif/` hasil Run all di Colab

## Jalankan di komputer sendiri (opsional)
```
pip install -r requirements.txt
streamlit run app.py
```

## Deploy gratis ke Streamlit Community Cloud
1. Buat repo publik di GitHub, upload semua isi folder ini **plus** `fact_dashboard.xlsx` (data agregat, bukan data mentah).
2. Buka share.streamlit.io, login dengan GitHub, klik **Create app**.
3. Pilih repo, branch `main`, main file path `app.py`, lalu **Deploy**.
4. Setelah jadi, salin link `https://....streamlit.app` dan tempel di deck (isi `DASHBOARD_URL` di notebook, jalankan ulang cell deck).
5. Buka link-nya sebentar sebelum dikirim. App gratis bisa "tidur" kalau lama tidak dibuka, dan butuh beberapa detik untuk bangun.

## Kalau data berubah
Jalankan ulang notebook, ganti `fact_dashboard.xlsx` di repo, commit. App otomatis reload.

## Catatan angka
- Total watch time memakai nilai asli; rata-rata durasi per play memakai nilai yang dibatasi 40 menit (`watch_seconds_capped`) supaya outlier tidak menggeser rata-rata.
- Chart pola jam, heatmap, dan weekend vs weekday tidak menghitung hari pertama dan terakhir kalau datanya terpotong zona waktu (otomatis terdeteksi dari volume plays).
- Semua angka berubah mengikuti filter di bagian atas.
