# panel-server
Server Node.js sederhana untuk menampilkan panel skrining antibodi dari Excel (antigram.xlsx).
- Taruh file Excel di `data/antigram.xlsx`
- Endpoint: GET /panel.json
- Frontend: static/panel_client.html

## Setup lokal
1. Install Node.js (v14+)
2. Clone repo, masuk folder
3. `npm install`
4. letakkan `antigram.xlsx` di folder `data/`
5. `npm start`
6. Buka http://localhost:3000/

## Deploy ke Render
1. Push repo ke GitHub.
2. Di Render.com -> New -> Web Service.
3. Connect GitHub repo, pilih branch (main).
4. Build Command: `npm install`
5. Start Command: `npm start`
6. (Optional) Set env var `UPLOAD_TOKEN` pada Render (Settings -> Environment) untuk mengamankan endpoint /upload.
7. Deploy.

## Mengganti file Excel (dua cara)
- **Rekomendasi (persisten):** Replace `data/antigram.xlsx` di repo dan push perubahan → redeploy (Render akan mengambil file baru).
- **Praktis (langsung):** Gunakan upload endpoint (hati-hati: butuh token jika di-set):
