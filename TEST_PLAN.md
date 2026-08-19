# Test Plan WordSplitter

Pengujian dibagi menjadi dua lapis. Lapis pertama adalah pengujian otomatis yang dapat dijalankan di mana saja tanpa Microsoft Word. Lapis kedua adalah pengujian manual pada Windows 11 dengan Word terpasang, karena perilaku *rendering* dan COM hanya dapat diverifikasi pada lingkungan sebenarnya.

---

## A. Pengujian otomatis

```bat
py -3 tests\test_core.py
py -3 tests\test_split_flow.py
```

### A.1 `test_core.py` (42 pengujian, seluruhnya lulus)

| Kelompok | Cakupan |
|---|---|
| `TestPageParsing` | Angka valid, input kosong, input non numerik (`abc`, `3.5`, `1e3`, `10a`, `--4`), nilai nol, nilai negatif |
| `TestSplitPlan` | Halaman berdampingan, jarak dua halaman, normalisasi urutan, batas pertama, batas terakhir, halaman identik, halaman di luar rentang, dokumen satu halaman, dan pemeriksaan menyeluruh atas seluruh kombinasi A dan B untuk dokumen 2 sampai 29 halaman guna memastikan tidak ada kombinasi yang menghasilkan bagian kosong |
| `TestOutputNames` | Penambahan ekstensi, ekstensi ganda, sembilan karakter terlarang Windows, nama kosong, nama yang hanya berisi ekstensi, nama cadangan (`CON`, `NUL`, `COM1`, `LPT9`), akhiran titik dan spasi, karakter kontrol, nama melebihi 180 karakter |
| `TestOutputPaths` | Resolusi ke direktori sumber, nama identik yang berbeda kapitalisasi, penimpaan file sumber, deteksi file keluaran yang sudah ada |
| `TestInputValidation` | Dokumen valid, file hilang, ekstensi salah, arsip rusak, file 0 byte, arsip ZIP tanpa `word/document.xml`, seleksi kosong |
| `TestUtils` | Format ukuran, normalisasi ekstensi, keunikan *temporary path* dan lokasinya pada *volume* tujuan, `safe_unlink` yang tidak pernah melempar, perbandingan biner, deteksi direktori yang tidak dapat ditulis, estimasi ruang disk |
| `TestModuleIntegrity` | Seluruh modul non GUI dapat diimpor, ketiadaan Word dilaporkan tanpa *exception*, permukaan API `split_engine` lengkap |

### A.2 `test_split_flow.py` (18 pengujian, seluruhnya lulus)

Sesi Word disubstitusi dengan *test double* deterministik yang memodelkan dokumen sebagai aliran karakter dengan 100 karakter per halaman.

| Skenario | Ekspektasi | Hasil |
|---|---|---|
| Analisis dokumen 10 halaman | Melaporkan 10 halaman | Lulus |
| Pemisahan pada Page A 5 dan Page B 6 | Part 1 dan Part 2 saling melengkapi, total karakter utuh, tanpa sisa file sementara | Lulus |
| Jarak dua halaman (4 dan 6) | `split_page` = 6 sesuai konvensi terdokumentasi | Lulus |
| Urutan terbalik (6 dan 5) | Hasil identik dengan urutan menaik | Lulus |
| Default tanpa penghapusan | File asli tetap ada | Lulus |
| `delete_original` aktif | File asli terhapus hanya setelah kedua keluaran tervalidasi | Lulus |
| File sumber hilang saat proses | `SplitError` yang informatif | Lulus |
| File sumber rusak | `SplitError`, proses tidak pernah menyentuh Word | Lulus |
| Keluaran sudah ada tanpa izin timpa | Proses ditolak, file lama tidak berubah | Lulus |
| Keluaran sudah ada dengan izin timpa | File lama tergantikan | Lulus |
| Kegagalan saat membentuk Part 2 | Tidak ada file final, file asli utuh, tidak ada sisa file sementara | Lulus |
| Kegagalan validasi keluaran dengan `delete_original` aktif | File asli **tetap ada**, tidak ada file final | Lulus |
| Penghapusan rentang gagal secara diam diam | Terdeteksi melalui jumlah halaman, proses dinyatakan gagal | Lulus |
| Keluaran identik secara biner | Diperlakukan sebagai peringatan, bukan kegagalan | Lulus |
| Halaman di luar rentang | Ditolak sebelum file apa pun dibuat | Lulus |
| Pembatalan pengguna | Berhenti sebelum *commit*, seluruh file sementara dibersihkan | Lulus |
| Pesan kemajuan | Emisi status untuk Part 1 dan Part 2 | Lulus |
| Lokasi *temporary file* | Berada pada *volume* yang sama dengan tujuan | Lulus |

### A.3 Pengujian antarmuka

Konstruksi jendela, transisi status *busy*, pengaktifan dan penonaktifan kontrol, penanganan pesan antrean, serta dialog konflik telah diverifikasi melalui *smoke test* Tkinter pada mode *headless*. Seluruh transisi berjalan tanpa *exception*.

---

## B. Pengujian manual pada Windows 11

Siapkan dokumen uji berikut, lalu jalankan aplikasi baik dari `src\main.py` maupun dari `dist\WordSplitter.exe`.

| No | Skenario | Prosedur | Kriteria lulus |
|---|---|---|---|
| 1 | Dokumen pendek | Dokumen 3 halaman teks biasa, Page A 1 dan Page B 2 | Part 1 berisi halaman 1, Part 2 berisi halaman 2 sampai 3 |
| 2 | Dokumen panjang | Dokumen 200 halaman, Page A 100 dan Page B 101 | Jumlah halaman kedua bagian menjumlah mendekati 200, antarmuka tetap responsif selama proses |
| 3 | Dokumen dengan tabel | Tabel melintasi titik pisah dan tabel utuh di kedua sisi | Tabel yang berada penuh di satu sisi tetap utuh beserta *border*, *shading*, dan *merged cell* |
| 4 | Dokumen dengan gambar | Gambar *inline* dan *floating* pada kedua sisi | Gambar tetap ada, ukuran dan posisi tidak berubah |
| 5 | Dokumen dengan formatting kompleks | Campuran *style*, *font*, warna, *highlight*, *superscript* | Seluruh format bertahan |
| 6 | Dokumen dengan section break | *Section* dengan orientasi dan margin berbeda | Bagian yang tidak dihapus mempertahankan *page setup* miliknya. Perilaku penggabungan *section* sesuai catatan keterbatasan pada README bagian 3.4 |
| 7 | Dokumen dengan page break manual | Titik pisah tepat pada *page break* | Tidak ada halaman kosong tambahan pada awal Part 2 |
| 8 | Halaman pertama | Page A 1 dan Page B 2 | Part 1 berisi tepat satu halaman, tidak ada bagian kosong |
| 9 | Halaman terakhir | Page A n-1 dan Page B n | Part 2 berisi tepat satu halaman |
| 10 | Input halaman invalid | `0`, `-3`, `abc`, `2.5`, kosong, dua nilai sama, nilai melebihi jumlah halaman | Dialog kesalahan yang jelas, aplikasi tidak *crash*, tidak ada file dibuat |
| 11 | File output sudah ada | Jalankan dua kali dengan nama sama | Dialog `File already exists.` muncul dengan tiga pilihan. `Cancel` tidak mengubah apa pun. `Choose Another Name` mengembalikan ke formulir. `Replace` menimpa |
| 12 | Permission denied | Simpan dokumen pada folder *read-only* atau `C:\Program Files` | Pesan izin yang jelas sebelum Word dijalankan, tidak ada *traceback* |
| 13 | Penghapusan file asli | Centang `Delete original file` | Dialog konfirmasi muncul. File asli hilang hanya setelah kedua keluaran ada dan tervalidasi |
| 14 | Kegagalan di tengah proses | Tutup paksa proses `WINWORD.EXE` melalui Task Manager saat status `Processing...` | Pesan kesalahan yang jelas, file asli utuh, tidak ada file `~wordsplitter_*.tmp.docx` tertinggal |
| 15 | Dokumen corrupt | Ubah beberapa byte pada file `.docx` menggunakan *hex editor* | Ditolak pada tahap validasi paket sebelum Word dijalankan |
| 16 | Nama file invalid | `Part<1`, `CON`, `Part 1.`, nama kosong, nama 300 karakter | Ditolak dengan pesan spesifik |
| 17 | File sedang digunakan | Buka dokumen di Word lalu tekan `Process` | Proses tetap berhasil karena pembacaan bersifat *read-only*. Bila Word mengunci secara eksklusif, muncul pesan yang meminta pengguna menutup dokumen |
| 18 | Validasi hasil | Buka kedua keluaran di Word | Keduanya terbuka tanpa dialog perbaikan, jumlah halaman sesuai rencana pemisahan |
| 19 | Word tidak terpasang | Jalankan pada mesin tanpa Word desktop | Pesan yang menjelaskan kebutuhan Microsoft Word, aplikasi tidak *crash* |
| 20 | Pembatalan | Tekan `Cancel` saat dokumen besar sedang diproses | Proses berhenti pada batas tahap berikutnya, tidak ada file final, tidak ada file sementara tertinggal |
| 21 | Dokumen terproteksi kata sandi | Pilih dokumen berkata sandi | Pesan kegagalan pembukaan yang jelas tanpa dialog Word yang menggantung |
| 22 | Disk penuh | Gunakan *volume* dengan sisa ruang di bawah empat kali ukuran dokumen | Ditolak sebelum proses dimulai dengan angka kebutuhan dan ketersediaan |
| 23 | Isi log | Periksa `%LOCALAPPDATA%\WordSplitter\logs\wordsplitter.log` | Memuat waktu, jalur file, ukuran, halaman terpilih, rencana pemisahan, dan hasil |
| 24 | Build executable | Jalankan `build.bat` pada mesin bersih | `dist\WordSplitter.exe` terbentuk dan berjalan tanpa instalasi Python |

---

## C. Cacat yang ditemukan dan diperbaiki selama pengembangan

| Temuan | Perbaikan |
|---|---|
| Pemeriksaan keluaran identik menggunakan perbandingan biner dapat menolak pemisahan yang sebenarnya benar ketika kedua bagian kebetulan memiliki isi sama | Kriteria kegagalan diganti menjadi pemeriksaan berbasis jumlah halaman, yaitu kegagalan hanya dinyatakan bila kedua bagian sama sama masih memuat seluruh dokumen. Kesamaan biner diturunkan menjadi peringatan |
| Rollback menghapus Part 1 yang sudah terpromosi meskipun nama tersebut sebelumnya ditempati file milik pengguna, sehingga menimbulkan kehilangan data | Pra eksistensi nama keluaran dicatat sebelum proses. File milik pengguna tidak pernah dihapus pada jalur rollback |
| *Lock file* Word berpola `~$` berpotensi tertinggal bila sesi Word berakhir tidak normal | Pembersihan file sementara diperluas untuk mencakup kedua varian penamaan *owner file* Word |
