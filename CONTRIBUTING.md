# Panduan Berpartisipasi di Sahabat Ibadah

Terima kasih sudah ingin membantu mengembangkan Sahabat Ibadah. Anda tidak harus menjadi programmer untuk berpartisipasi. Anda bisa melaporkan masalah, mengusulkan fitur, memperbaiki panduan, atau mengirim perubahan code.

Repositori proyek: [stenlymo/sabahat_ibadah](https://github.com/stenlymo/sabahat_ibadah)

## Pilih cara berpartisipasi

### 1. Melaporkan masalah atau mengusulkan fitur

Cara ini cocok jika Anda pengguna aplikasi dan belum ingin mengubah code.

1. Buka halaman [Issues](https://github.com/stenlymo/sabahat_ibadah/issues) proyek.
2. Klik **New issue**.
3. Beri judul singkat, misalnya `Tombol Berikutnya tidak berpindah slide` atau `Usulan: pilihan ukuran huruf`.
4. Jelaskan langkah untuk melihat masalah, apa yang terjadi, dan apa yang Anda harapkan.
5. Jika membantu, lampirkan tangkapan layar yang tidak memuat informasi pribadi.

Untuk usulan fitur, jelaskan kebutuhan operator dan contoh cara fitur tersebut akan digunakan. Periksa Issues yang sudah ada agar laporan tidak terduplikasi.

### 2. Membantu tanpa mengubah program

Anda dapat membantu dengan:

- Memperjelas README atau panduan penggunaan.
- Menunjukkan langkah yang membingungkan bagi operator baru.
- Menguji alur aplikasi dan melaporkan hasilnya melalui Issue.
- Mengusulkan perbaikan teks tombol atau istilah agar lebih mudah dimengerti.

Untuk laporan, sertakan versi Windows, cara menjalankan aplikasi, tab/menu yang digunakan, dan urutan klik yang menimbulkan masalah. Jangan sertakan kata sandi, token, database pribadi, rekaman kamera, atau data ibadah yang tidak boleh dibagikan.

### 3. Mengirim perubahan code dengan GitHub Desktop

GitHub Desktop menyediakan antarmuka grafis sehingga Anda tidak perlu mengetik perintah Git.

#### A. Buat salinan proyek (fork)

1. Masuk ke akun GitHub Anda.
2. Buka halaman [repositori Sahabat Ibadah](https://github.com/stenlymo/sabahat_ibadah).
3. Klik **Fork**, lalu buat fork di akun Anda.
4. Buka fork yang baru dibuat dan salin alamatnya melalui **Code → HTTPS**.

#### B. Siapkan salinan di komputer

1. Pasang [GitHub Desktop](https://desktop.github.com/), lalu masuk dengan akun GitHub.
2. Di GitHub Desktop pilih **File → Clone repository… → URL**.
3. Tempel alamat HTTPS fork Anda, pilih folder tujuan, lalu klik **Clone**.
4. Pilih **Current branch → New branch**, beri nama yang menjelaskan pekerjaan, misalnya `perbaiki-navigasi-slide`.
5. Jika ingin menjalankan aplikasi dari source, pasang Python 3.10 atau lebih baru. Buka PowerShell di folder clone dan jalankan:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   python -m pip install -r requirements.txt
   python .\sahabat_ibadah.py
   ```

Fitur tertentu memerlukan perangkat lunak atau perangkat tambahan. Lihat bagian **Menjalankan aplikasi**, **Kamera & Live**, dan **Broadcast** di [README](README.md).

#### C. Buat perubahan dan ajukan

1. Ubah file yang terkait pekerjaan Anda. Buat perubahan kecil dan fokus pada satu perbaikan atau fitur.
2. Jalankan aplikasi dan periksa perubahan yang Anda buat. Untuk perubahan kamera, mikrofon, proyektor, atau OBS, uji dengan perangkat terkait jika tersedia.
3. Di GitHub Desktop, tinjau daftar perubahan. Isi ringkasan commit, misalnya `Perbaiki perpindahan slide Broadcast`, lalu klik **Commit to ...**.
4. Klik **Push origin** untuk mengirim branch ke fork Anda.
5. Klik **Create Pull Request**. Pastikan tujuan Pull Request adalah repositori `stenlymo/sabahat_ibadah` pada branch `main`.
6. Jelaskan masalah yang diperbaiki, perubahan yang dibuat, serta cara Anda memeriksanya. Klik **Create pull request**.
7. Tanggapi pertanyaan atau masukan pada Pull Request. Pemilik proyek akan meninjau dan menentukan apakah perubahan dapat digabungkan.

Panduan resmi: [Fork a repository](https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/working-with-forks/fork-a-repo) dan [Contributing to open source](https://docs.github.com/en/get-started/exploring-projects-on-github/contributing-to-open-source).

## Jika Anda sudah mendapat akses langsung

Pemilik proyek dapat mengundang Anda melalui **Settings → Collaborators → Add people**. Setelah menerima undangan, Anda dapat bekerja dari repositori utama. Untuk memudahkan peninjauan, tetap buat branch baru dan ajukan Pull Request; hindari mengubah branch `main` secara langsung.

## Lindungi data dan file pribadi

Jangan mengirim file berikut ke GitHub:

- Database aplikasi seperti `sahabat_ibadah.db`, yang mungkin berisi lagu, jadwal, atau data lokal Anda.
- Folder `.venv`, cache, hasil build, cadangan ZIP, dan file aplikasi `.exe`.
- Dataset Alkitab atau materi lain yang lisensinya tidak mengizinkan distribusi ulang.
- Video kamera, foto pribadi, atau rekaman ibadah tanpa izin untuk membagikannya.

Repositori menyediakan `.gitignore` untuk mengecualikan sebagian besar file tersebut. Tetap periksa daftar perubahan sebelum commit. Jika tidak sengaja mengunggah rahasia atau data pribadi, segera beri tahu pemilik proyek; menghapus file pada commit berikutnya belum tentu menghapusnya dari riwayat Git.

## Yang diharapkan dari Pull Request

- Gunakan bahasa yang jelas dan sopan.
- Sebutkan Issue terkait jika ada.
- Jelaskan perubahan dan cara memeriksanya.
- Hindari menggabungkan banyak perubahan yang tidak berhubungan dalam satu Pull Request.
- Jangan mengunggah file build atau database pengguna.

Terima kasih sudah membantu Sahabat Ibadah menjadi lebih bermanfaat bagi operator dan jemaat.
