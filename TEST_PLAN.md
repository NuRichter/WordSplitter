# Test Plan WordSplitter 2.0

Pengujian dibagi menjadi tiga lapis.

1. **Pengujian otomatis** di `tests/`. Sebagian besar berjalan di mana saja; pengujian yang memakai Microsoft Word asli dilewati otomatis bila Word tidak ada.
2. **Laboratorium skenario** dengan Microsoft Word asli, yang dipakai untuk membuktikan cacat versi 1.0 dan menguji metode potong versi 2.0 sebelum diterapkan.
3. **Pemeriksaan manual** pada Windows untuk hal yang hanya bisa dinilai manusia: tampilan, bahasa, dan rasa pakai.

---

## A. Pengujian otomatis

```bat
py -3 -m unittest discover -s tests -v
set WS_SKIP_WORD=1 & py -3 -m unittest discover -s tests     (tanpa Word)
```

`build.bat` menjalankan seluruh pengujian dan menghentikan build bila ada yang gagal. Hasil terakhir: **93 pengujian, semuanya lulus**, termasuk dua pengujian Word asli. Suite juga lulus dengan `-W error::ResourceWarning`.

| Berkas | Cakupan utama |
|---|---|
| `test_core.py` | Angka halaman (termasuk digit Arab, Persia, Devanagari, dan angka lebar); rencana potong untuk **semua** pasangan A/B pada dokumen 2–39 halaman tanpa bagian kosong; aturan nama file Windows (karakter terlarang, nama cadangan termasuk `con.backup`, akhiran titik/spasi, karakter kontrol, panjang); path terlalu panjang untuk Word; pemeriksaan paket `.docx` (kosong, bukan ZIP, terenkripsi/`.doc` lama, entri hilang, arsip terpotong, CRC rusak); salin atomik tanpa sisa file; sidik jari file |
| `test_engine.py` | Alur dengan sesi Word simulasi: analisis tidak pernah membuka file asli; File 1 + File 2 = dokumen asli; urutan A/B terbalik; aturan titik tengah; tanpa sisa file sementara; perubahan file setelah analisis; pergeseran jumlah halaman; file tujuan sudah ada dengan dan tanpa izin timpa; kegagalan membentuk Part 2; verifikasi gagal saat opsi hapus aktif (file asli tetap ada); penghapusan diam-diam yang gagal; **kegagalan commit di tengah dengan pemulihan file lama**; hapus permanen di drive tanpa Recycle Bin; Recycle Bin berhasil dan gagal; peringatan pergeseran tata letak dan *field* yang diubah; pembatalan; urutan kejadian kemajuan |
| `test_i18n.py` | 30 bahasa dengan berkas lengkap; kunci sama persis; *placeholder* sama dengan bahasa Inggris; semua string bisa dirender; string RTL tanpa tanda kurung dan tanpa rentang angka bertanda hubung; nama default dan nama sampel valid sebagai nama file; label diterjemahkan di dalam pesan; format ukuran per bahasa; pemetaan tag bahasa Windows |
| `test_gui.py` | Jendela dibangun dalam 30 bahasa × 2 tema; pratinjau mengikuti isian halaman; status kontrol; ganti bahasa mempertahankan nama yang diketik pengguna; tema tersimpan; **tidak ada timer saat idle**; **proses jendela tidak memuat pywin32/COM**; kesalahan input tampil sebagai dialog; mode ringkas muat di layar 1366×768 |
| `test_worker.py` | Protokol worker lewat proses sungguhan; dokumen tidak valid ditolak sebelum Word dijalankan; dengan Word asli: dokumen latihan tepat 12 halaman dan terbelah 6/6; pembatalan di tengah pemisahan tidak meninggalkan file atau proses Word |

---

## B. Laboratorium skenario dengan Microsoft Word asli

Dokumen uji dibuat langsung oleh Word: teks mengalir (A), *page break* manual (B), tiga *section* portrait/landscape/portrait dengan *header* dan penomoran berbeda (C), tabel 150 baris dengan baris judul berulang (D), *Track Changes* aktif (E), daftar isi yang membentang dua halaman (H), dan dokumen besar 289 halaman.

### B.1 Bukti cacat versi 1.0 dan hasil versi 2.0

Kolom "File 1/File 2" berisi jumlah halaman hasil dibanding jumlah yang seharusnya.

| Dokumen | A\|B | v1.0 | v2.0 |
|---|---|---|---|
| A teks mengalir | 6\|7 | 6/7 benar | 6/7 benar |
| B *page break* | 6\|7 | **7**/6: halaman kosong di akhir File 1 | 6/6 |
| B *page break* | 11\|12 | **12**/1 | 11/1 |
| B *page break* | 1\|2 | – | 1/11 |
| C *section* | 1\|2 | **2**/6, File 1 ber-*header* "BODY" | 1/6, *header* "FRONT", nomor romawi |
| C *section* | 2\|3 | **3**/5, ada *section* kosong ber-*header* salah | 2/5 |
| C *section* | 3\|4 | – | 3/4, *section* terakhir File 1 tetap *landscape* |
| C *section* | 4\|5 | **5**/3 | 4/3 |
| C *section* | 5\|6 | – | 5/2 |
| D tabel | 3\|4 | **4**/3, baris judul hilang di File 2 | 3/3, baris judul ikut |
| D tabel | 1\|2 | – | 1/5, baris judul ikut |
| E *track changes* | 3\|4 | **4**/3, isi File 2 tersimpan sebagai revisi di File 1 | 3/3, 0 revisi bocor, setelan tetap aktif |
| H daftar isi | 1\|2 | **gagal** `Cannot edit Range` | 1/11, daftar isi diubah jadi teks, ada peringatan |
| H daftar isi | 5\|6 | – | 5/7 |

### B.2 Skenario README (dokumen di `samples/`)

| Dokumen | A\|B | Hasil v2.0 | Pemeriksaan |
|---|---|---|---|
| 01 dasar | 6\|7 | 6/6 | – |
| 01 dasar | 6\|8 | 7/5 | aturan titik tengah |
| 02 skripsi | 3\|4 | 3/16 | romawi + "Bagian Awal" di File 1, "BAB 1" + "Isi Skripsi" di File 2 |
| 02 skripsi | 9\|10 | 9/10 | *header* bab di File 2 |
| 02 skripsi | 13\|14 | 13/6 | File 2 *landscape* |
| 02 skripsi | 16\|17 | 16/3 | File 2 dibuka baris judul tabel, tanpa paragraf kosong |
| 03 lacak perubahan | 4\|5 | 4/4 | 2 revisi asli tetap di File 1, setelan aktif |

### B.3 Ketahanan

| Skenario | Hasil |
|---|---|
| Dokumen rusak (XML terpotong) | Pesan `err.word_open` dalam 2 detik, tanpa *hang*, tanpa sisa Word |
| Batal di tengah pemisahan dokumen 289 halaman | Berhenti dalam 0,04 detik; tidak ada file hasil, folder kerja, atau Word tersisa |
| Batal tepat saat WINWORD.EXE baru muncul | Berhenti dalam 0,53 detik; Word tidak tertinggal |
| Proses induk dibunuh saat worker berjalan | Worker menghentikan Word-nya sendiri; 0 Word tersisa |
| Dokumen sedang terbuka dan diedit di Word lain | Analisis dan pemisahan berhasil |
| Dialog modal Word ("File In Use", "caused a serious error") | Dideteksi *watchdog*, Word ditutup, pesan berisi judul dialog |
| Exe *onefile* memanggil dirinya sebagai worker | Satu folder ekstraksi dipakai bersama, kemajuan pertama 0,07 detik setelah spawn |

### B.4 Efisiensi

| Pengukuran | v1.0 | v2.0 |
|---|---|---|
| Siklus CPU idle per 15 detik (median 5) | 65,1 juta | 2,4 juta |
| Memori pribadi jendela | 22,4 MB | 24,0 MB |
| Waktu analisis dokumen contoh | – | 0,9–2,5 detik |
| Waktu pemisahan dokumen contoh | 1,2–3,9 detik (dengan hasil salah) | 1,3–4,3 detik |

---

## C. Pemeriksaan manual

| No | Skenario | Kriteria lulus |
|---|---|---|
| 1 | Buka Panduan, tekan *Buat contoh dan coba*, tekan Pisahkan | Dokumen latihan muncul di `Documents\WordSplitter`, Halaman A/B terisi 6/7, hasil 6 dan 6 halaman |
| 2 | Ganti ke setiap bahasa | Tidak ada teks terpotong, tidak ada kotak kosong pengganti huruf, istilah wajar untuk penutur asli |
| 3 | Bahasa Arab, Persia, Urdu | Tata letak cermin, rentang halaman terbaca benar, penanda daftar di tepi kanan |
| 4 | Tema terang dan gelap | Semua kontrol terbaca; *title bar* ikut gelap |
| 5 | Layar 1366×768 pada 100% | Jendela muat utuh dalam mode ringkas |
| 6 | Skala tampilan 125% dan 150% | Teks tajam, tidak ada elemen terpotong |
| 7 | Input halaman salah (`0`, `-3`, `abc`, kosong, A = B, melebihi jumlah halaman) | Pratinjau menampilkan pesan merah; tombol Pisahkan memunculkan dialog yang jelas |
| 8 | Nama file sudah ada | Dialog *Timpa* / *Pakai nama lain* / *Batal* bekerja sesuai label |
| 9 | Hapus file asli di drive C: | File asli masuk Recycle Bin |
| 10 | Tutup jendela saat proses berjalan | Dialog konfirmasi; setelah ditutup tidak ada WINWORD.EXE tersisa |
| 11 | Folder *read-only* atau dilindungi *Controlled folder access* | Pesan jelas sebelum Word dijalankan |
| 12 | Komputer tanpa Word | Pesan yang menjelaskan kebutuhan Word |
| 13 | `build.bat` di mesin bersih | Semua pengujian lulus, `dist\WordSplitter.exe` terbentuk dan berjalan tanpa Python |

---

## D. Temuan selama pengembangan versi 2.0

| Temuan | Perbaikan |
|---|---|
| Lima cacat akurasi v1.0 pada tabel B.1 | Algoritma potong baru (bagian 6 README), dibuktikan per kasus |
| Word yang dibunuh saat dokumen terbuka memicu dialog pemulihan "caused a serious error" pada pembukaan berikutnya, sehingga otomasi *hang* | Word hanya membuka salinan kerja bernama unik yang tidak pernah dipakai ulang; *watchdog* menangani dialog yang tetap muncul |
| File yang dipegang instance Word lain memicu dialog "File In Use" meski `DisplayAlerts` mati | Snapshot ke folder kerja pribadi |
| Worker yang dibunuh saat Word sedang mulai meninggalkan WINWORD.EXE yatim | Pembatalan dua fase lewat stdin: worker menunggu PID Word diketahui lalu menghentikannya sendiri; penanda batal global agar sesi yang baru mulai pun ikut batal |
| Menunggu WINWORD.EXE keluar setelah `Quit` memakan 2,6 detik per tugas | Hasil dikirim dulu, penutupan Word ditunggu di latar belakang |
| *Watchdog* sempat menganggap jendela utama Word yang muncul sebagai dialog | Jendela utama disembunyikan lagi; hanya jendela dialog yang memicu penghentian |
| Penghapusan paragraf kosong di ujung File 1 ditolak Word bila paragraf itu menutup *field* daftar isi | Paragraf berisi *field* atau gambar tidak dianggap kosong; perapian ujung bersifat kosmetik dan tidak pernah menggagalkan proses |
| Pola `kondisi and 0 or 2` di tata letak RTL selalu menghasilkan 2 | Diganti ekspresi kondisional eksplisit; tertangkap oleh tangkapan layar bahasa Arab |
| Tk membalik rentang angka "1-24" di kalimat RTL | Rentang ditulis dengan kata ("من الصفحة 1 إلى 24", "1 تا 24") dan dijaga oleh `test_i18n.py` |
| `PageSetup` yang di-*assign* utuh lewat COM diam-diam tidak berpengaruh | Properti disalin satu per satu |
