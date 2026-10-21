# Panduan Sahabat Ibadah

Sahabat Ibadah adalah aplikasi desktop Windows untuk menyiapkan susunan ibadah, mengelola lagu dan ayat Alkitab, serta menampilkan kamera atau materi ibadah ke proyektor dan aplikasi rapat daring. Data aplikasi disimpan secara lokal.

Ingin melaporkan masalah atau ikut mengembangkan aplikasi? Mulai dari [Panduan Berpartisipasi](CONTRIBUTING.md).

## Daftar isi

- [Menjalankan aplikasi](#menjalankan-aplikasi)
- [Alur singkat ibadah](#alur-singkat-ibadah)
- [Susunan Ibadah dan Resources](#susunan-ibadah-dan-resources)
- [Perpustakaan Lagu](#perpustakaan-lagu)
- [Ayat Alkitab](#ayat-alkitab)
- [Proyeksi ibadah](#proyeksi-ibadah)
- [Kamera & Live](#kamera--live)
- [Broadcast ke Meet, Zoom, atau Teams](#broadcast-ke-meet-zoom-atau-teams)
- [Tema dan warna](#tema-dan-warna)
- [Cadangan dan pemulihan data](#cadangan-dan-pemulihan-data)
- [Pemecahan masalah](#pemecahan-masalah)

## Menjalankan aplikasi

### Menjalankan dari kode sumber

1. Pasang Python 3.10 atau yang lebih baru.
2. Buka PowerShell di folder aplikasi.
3. Buat dan aktifkan lingkungan Python jika belum tersedia:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   ```

4. Jalankan aplikasi:

   ```powershell
   python .\sahabat_ibadah.py
   ```

Dependensi dalam `requirements.txt` dipakai untuk kamera, video, mikrofon, dan kamera virtual. Pada Windows, fitur presentasi PowerPoint memerlukan Microsoft PowerPoint desktop. Fitur Broadcast juga memerlukan OBS Studio dengan OBS Virtual Camera yang tersedia.

Database lokal bernama `sahabat_ibadah.db`. Jangan hapus atau pindahkan file tersebut jika ingin mempertahankan data.

## Alur singkat ibadah

1. Pilih atau buat jadwal di **Susunan Ibadah**.
2. Tambahkan lagu dari **Perpustakaan Lagu**, ayat dari **Ayat Alkitab**, atau gambar/video/PowerPoint dari **Resources**.
3. Atur urutan item dan periksa pratinjau.
4. Klik **Mulai Proyeksi** untuk menampilkan susunan pada layar proyektor.
5. Jika ingin mengirim keluaran ke Meet, Zoom, atau Teams, buka **Broadcast** dan mulai output virtual camera.
6. Buat cadangan dari **Data & Backup** sebelum pemindahan atau perubahan besar.

## Susunan Ibadah dan Resources

**Susunan Ibadah** menyimpan jadwal ibadah dan urutan item yang akan ditampilkan. Gunakan tombol tambah untuk memasukkan lagu, ayat, atau media; tombol panah atas/bawah untuk mengubah urutan; dan **Edit** untuk memperbarui item. **Hapus dari Susunan** hanya menghapus item dari jadwal, bukan dari perpustakaan.

Tab **Resources** membantu mencari lagu, ayat, gambar, video, dan file PowerPoint lalu menambahkannya ke jadwal. File media ditautkan dari lokasi asalnya, jadi jangan pindahkan atau hapus file setelah ditambahkan.

## Perpustakaan Lagu

- Pilih lagu di daftar untuk melihat atau mengubah judul, penulis, dan lirik.
- **Lagu Baru** membuat entri lagu. Simpan perubahan dengan **Simpan Lagu**.
- Pisahkan bagian lagu dengan satu baris kosong. Setiap bagian menjadi slide tersendiri.
- Gunakan **Impor CSV** atau **Ekspor CSV** untuk memindahkan koleksi. CSV menggunakan kolom `title`, `author`, dan `lyrics`; kolom penulis boleh kosong.

## Ayat Alkitab

Tab **Ayat Alkitab** memiliki pencarian ayat dan penjelajahan menurut versi, kitab, pasal, serta ayat. Pilih versi yang sudah diimpor, lalu masukkan referensi atau cari kata dalam teks. Untuk menambahkan beberapa ayat, pilih baris dengan Ctrl atau Shift lalu klik **Tambah Ayat Terpilih**. Rentang seperti `Kejadian 1:4-9` juga dapat ditambahkan.

Untuk mengimpor terjemahan, klik **Impor Dataset Alkitab CSV…**. Dataset AYT dan dataset multi-versi seperti KJV, AMP, atau NIV memakai format CSV yang didukung aplikasi. Perhatikan lisensi dan atribusi dari penyedia terjemahan, terutama saat mendistribusikan data atau materi turunannya.

## Proyeksi ibadah

Klik **Mulai Proyeksi** untuk membuka tampilan layar penuh. Jika ada monitor kedua yang tersambung dalam mode extended, aplikasi akan menggunakannya untuk proyeksi dan mempertahankan kontrol pada monitor operator.

Kontrol papan ketik pada jendela proyeksi:

| Tombol | Fungsi |
| --- | --- |
| Panah kiri / kanan atau spasi | Pindah ke bagian sebelumnya / berikutnya |
| Angka lalu Enter | Lompat ke bagian lagu atau ayat |
| Home / End | Ke bagian awal / akhir |
| B | Layar hitam (blackout) |
| C | Bersihkan tampilan |
| Esc | Tutup proyeksi |

Tombol **Blackout** dan **Clear** juga tersedia pada kontrol proyeksi.

## Kamera & Live

Tab **Kamera & Live** dipakai untuk menampilkan webcam bawaan, webcam USB, USB capture device, atau file video. Klik **Cari** untuk memperbarui daftar perangkat. Pilih kamera utama dan, bila diperlukan, kamera kedua; pilih layar proyeksi; lalu klik **Mulai Kamera / Video**.

Preset kamera yang tersedia:

- **Kamera 1 penuh** atau **Kamera 2 penuh**
- **Dua kamera 50:50**
- **Sembunyikan kamera**

Teks lagu atau ayat dapat ditampilkan sebagai overlay. Kontrolnya mencakup warna teks, garis tepi, posisi, perataan, dan margin. Gunakan **Atur Posisi & Margin…** untuk menata posisi teks. Untuk presentasi, klik **Pilih Media…** lalu gunakan kontrol slide. Media dapat ditampilkan bersama kamera, layar penuh, atau disembunyikan kembali.

## Broadcast ke Meet, Zoom, atau Teams

Broadcast membuat perangkat **OBS Virtual Camera** yang dapat dipilih sebagai kamera di aplikasi rapat. Video dikirim oleh Sahabat Ibadah; suara rapat tetap menggunakan mikrofon yang dipilih langsung di Meet, Zoom, atau Teams. Jika memakai mikrofon dari aplikasi rapat, pilih perangkat mikrofon yang sama di pengaturan aplikasi rapat.

### Menyiapkan siaran

1. Pastikan OBS Studio terpasang dan OBS Virtual Camera tersedia. Jika perlu, buka OBS dan aktifkan **Start Virtual Camera**.
2. Buka tab **Broadcast**.
3. Di **Sumber output**, pilih salah satu:
   - **Kamera** untuk siaran kamera.
   - **Item Susunan** untuk menyiarkan lagu, ayat, gambar, atau slide PowerPoint dari jadwal.
   - **Kamera + Item Susunan** untuk menggabungkan kamera dengan materi jadwal.
4. Pilih item susunan. Untuk lagu, tata letak **Lagu Overlay Kamera** menampilkan lirik di atas kamera. Ayat dan media dapat ditampilkan layar penuh atau berdampingan dengan kamera.
5. Pilih perangkat kamera dan **Tampilan output** yang sesuai. Pilihan tata letak berubah sesuai jenis sumber.
6. Klik **Mulai Output Kamera / Konten**. Pratinjau Broadcast memperlihatkan komposisi yang dikirim ke perangkat virtual camera.
7. Di pengaturan kamera Meet, Zoom, atau Teams, pilih **OBS Virtual Camera**.
8. Untuk lagu, ayat, atau PowerPoint yang memiliki beberapa slide, gunakan **Sebelumnya** dan **Berikutnya**.

Output virtual camera menggunakan 1280 × 720 pada 25 FPS. Pratinjau operator diperbarui lebih jarang agar aplikasi tetap ringan.

### Mengubah tampilan saat siaran aktif

Sumber dan tata letak dapat diubah dari kontrol Broadcast selama output aktif. Pilih item atau tata letak baru untuk memperbarui keluaran. Perubahan teks seperti warna, posisi, margin, dan pembalikan horizontal berlaku pada komposisi yang sedang ditampilkan.

### Menghentikan sementara atau mengakhiri siaran

- **Hentikan Sumber · Tampilkan Hitam** menghentikan gambar sumber dan mengirim layar hitam sementara output virtual camera tetap tersambung. Klik **Mulai Kembali Sumber** untuk membuka kembali kamera atau materi yang terakhir dipilih.
- **Hentikan Output** mengakhiri koneksi virtual camera. Mulai lagi dengan **Mulai Output Kamera / Konten**.
- Jika kamera tidak dapat dibuka kembali, pastikan kamera tidak sedang dipakai aplikasi lain, lalu klik **Coba Lagi Aktifkan Sumber** atau mulai ulang output.

Di bagian **Posisi dan arah teks**, pilih warna teks lagu atau ayat secara terpisah, perataan horizontal dan vertikal, margin, serta opsi **Balik horizontal seluruh output** bila memang diperlukan. Roda mouse di atas dropdown Broadcast menggulir panel; untuk mengubah pilihan, klik dropdown lalu klik opsi.

## Tema dan warna

- Tombol **Mode Terang / Mode Gelap** mengganti tema aplikasi.
- **List Warna** membuka palet untuk mengatur warna antarmuka pada kedua mode. Perubahan disimpan otomatis.
- **Reset Mode Gelap** dan **Reset Mode Terang** mengembalikan palet masing-masing ke warna bawaan.
- Pemilih warna teks di Broadcast khusus untuk warna teks lagu dan ayat; warna keduanya dapat diatur terpisah.

## Cadangan dan pemulihan data

1. Klik **Data & Backup**.
2. Pilih **Buat Cadangan…** lalu tentukan lokasi file ZIP.
3. Untuk mengembalikan data, pilih **Pulihkan Cadangan…** dan pilih file ZIP Sahabat Ibadah.

Pemulihan mengganti data lokal yang sedang digunakan dengan isi cadangan. Pastikan cadangan yang benar dipilih. Simpan cadangan terbaru sebelum memulihkan.

## Pemecahan masalah

### Kamera tidak muncul atau tidak dapat dibuka

- Klik **Cari** untuk memperbarui daftar kamera.
- Periksa izin kamera Windows dan tutup aplikasi lain yang mungkin sedang memakai kamera.
- Coba pilih perangkat lain atau mulai ulang aplikasi.
- Jika kamera yang sama sedang digunakan pada **Kamera & Live**, hindari memilih kamera tersebut dua kali sebagai dua input yang berbeda.

### OBS Virtual Camera tidak tersedia

- Pastikan OBS Studio terpasang dan OBS Virtual Camera muncul pada perangkat kamera Windows.
- Buka OBS dan klik **Start Virtual Camera** bila belum aktif, kemudian coba mulai output Broadcast kembali.
- Setelah output dimulai, pilih **OBS Virtual Camera** pada pengaturan kamera aplikasi rapat.

### PowerPoint tidak tampil

- Pastikan Microsoft PowerPoint desktop terpasang.
- File `.pptx`/`.ppt` harus tetap berada di lokasi yang sama. Animasi, transisi, audio, dan video tertanam tidak diputar oleh proses render slide.

### Panel atau kontrol tidak terlihat

Gunakan roda mouse atau scrollbar untuk menggulir panel Broadcast dan Kamera & Live. Pada dropdown Broadcast, roda mouse menggulir panel dan tidak mengganti pilihan.

### Aplikasi tidak responsif

Tutup output yang tidak digunakan, kurangi jumlah kamera/video yang aktif, dan coba siaran dengan satu kamera terlebih dahulu. Output Broadcast berjalan pada 25 FPS, sedangkan preview operator diperbarui sekitar dua kali per detik. Jika masalah berulang, catat apakah macet terjadi saat memilih kamera, saat output mulai, atau setelah berjalan beberapa waktu.

## Dataset AYT

Unduh `ayt.csv` dari [repositori resmi AYT](https://github.com/sabdacode/ayt/tree/main/csv), lalu impor melalui **Ayat Alkitab → Impor Dataset Alkitab CSV…**. Dataset AYT memiliki hak cipta YLSA-AYT dan ketentuan penggunaan/derivatif BY-NC-SA. Cantumkan atribusi dan ikuti ketentuan penggunaan. Lihat [informasi lisensi AYT](https://github.com/sabdacode/ayt#lisensi).
