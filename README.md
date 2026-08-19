# WordSplitter

Aplikasi *desktop* Windows 11 untuk memisahkan satu dokumen Microsoft Word `.docx` menjadi dua dokumen berdasarkan posisi tengah antara dua halaman yang dipilih pengguna.

Aplikasi ini hanya memiliki satu fungsi. Tidak ada konversi format, tidak ada penyuntingan isi, tidak ada fitur tambahan.

```
Select File  →  Analyze  →  Page A & Page B  →  Process  →  Save
```

---

## 1. Persyaratan sistem

| Komponen | Persyaratan |
|---|---|
| Sistem operasi | Windows 11 64-bit (Windows 10 64-bit juga berfungsi, tetapi tidak diuji) |
| Microsoft Word | **Wajib.** Word desktop 2016, 2019, 2021, atau Microsoft 365, terpasang dan berlisensi aktif |
| Python | Hanya untuk *developer*. Versi 3.11 sampai 3.13. Pengguna akhir `.exe` tidak memerlukan Python |
| Dependency | `pywin32` untuk *runtime*, `pyinstaller` untuk *build* |

Microsoft Word versi *web* dan Microsoft Store *app* (UWP) tidak menyediakan antarmuka COM sehingga tidak dapat digunakan. Diperlukan instalasi Word desktop biasa (*Click-to-Run* atau MSI).

---

## 2. Mengapa Microsoft Word diperlukan

Format `.docx` menyimpan **aliran konten**, bukan halaman. Tidak ada satu pun elemen dalam `word/document.xml` yang menyatakan "di sini halaman 10 berakhir". Batas halaman dihitung pada saat *layout* oleh *rendering engine*, dan hasilnya bergantung pada metrik *font* yang terpasang, *printer driver* default, opsi kompatibilitas dokumen, serta versi Word itu sendiri.

Konsekuensinya, tiga pendekatan alternatif yang lazim dipakai semuanya ditolak dalam desain ini.

| Pendekatan | Alasan penolakan |
|---|---|
| `python-docx` dengan estimasi jumlah paragraf atau jumlah karakter per halaman | Tidak akurat. Kesalahan bertambah pada dokumen dengan tabel, gambar, *multi-column layout*, atau ukuran *font* campuran. Tidak ada jaminan posisi halaman sama sekali |
| Menghitung `w:lastRenderedPageBreak` di dalam XML | Penanda ini hanya ditulis oleh Word saat terakhir kali dokumen dirender dan tidak dijamin ada, tidak dijamin mutakhir, dan hilang pada dokumen yang dibuat oleh *generator* pihak ketiga |
| LibreOffice *headless* sebagai mesin pagination | Pagination LibreOffice berbeda dari Word pada dokumen kompleks. Konversi bolak balik juga berisiko mengubah *styles*, *numbering*, dan *content control* |

Karena itu WordSplitter menggunakan **Microsoft Word COM automation** melalui `pywin32`. Word adalah satu satunya otoritas yang benar untuk pertanyaan "di karakter ke berapa halaman N dimulai" pada mesin Windows 11 tersebut.

---

## 3. Pendekatan teknis pagination dan split

### 3.1 Penentuan titik pisah

1. Dokumen dibuka oleh instance Word tersendiri yang dibuat dengan `DispatchEx`, dalam mode `ReadOnly`, tidak terlihat, dan dengan `DisplayAlerts` dimatikan. Instance tersendiri memastikan jendela Word milik pengguna tidak terganggu dan tidak ikut tertutup.
2. `Document.Repaginate()` dipanggil agar tata letak mutakhir.
3. Jumlah halaman dibaca dari `Document.ComputeStatistics(wdStatisticPages)`.
4. Posisi awal halaman diperoleh dari `Document.GoTo(What=wdGoToPage, Which=wdGoToAbsolute, Count=N).Start`. Nilai ini adalah *character offset* pada aliran konten dokumen, dipetakan langsung oleh mesin *layout* Word.

### 3.2 Definisi "tengah antara Page A dan Page B"

Titik pisah berada pada batas halaman yang paling dekat dengan titik tengah Page A dan Page B. Aturan formalnya adalah

```
split_page = floor((A + B) / 2) + 1
Part 1 = halaman 1 .. split_page - 1
Part 2 = halaman split_page .. halaman terakhir
```

| Page A | Page B | Titik tengah | split_page | Hasil |
|---|---|---|---|---|
| 10 | 11 | 10,5 | 11 | Part 1 = 1-10, Part 2 = 11-akhir |
| 10 | 12 | 11 | 12 | Part 1 = 1-11, Part 2 = 12-akhir |
| 10 | 13 | 11,5 | 12 | Part 1 = 1-11, Part 2 = 12-akhir |
| 1 | 2 | 1,5 | 2 | Part 1 = 1, Part 2 = 2-akhir |

Ketika titik tengah jatuh tepat pada sebuah halaman, seperti pada kasus A = 10 dan B = 12, halaman tengah tersebut dimasukkan ke Part 1. Konvensi ini dipilih karena deterministik dan konsisten untuk semua kombinasi, dan diterapkan sama persis pada setiap proses. Urutan Page A dan Page B dinormalisasi secara otomatis sehingga 11 dan 10 menghasilkan titik pisah yang sama dengan 10 dan 11.

### 3.3 Metode pemisahan

Dokumen **tidak pernah dibangun ulang dari nol**. Prosedurnya adalah sebagai berikut.

1. File asli disalin dua kali dengan `shutil.copy2` ke dua file sementara pada direktori yang sama.
2. Salinan pertama dibuka oleh Word, lalu rentang karakter dari titik pisah sampai akhir dokumen dihapus. Hasilnya adalah Part 1.
3. Salinan kedua dibuka oleh Word, lalu rentang karakter dari awal dokumen sampai titik pisah dihapus. Hasilnya adalah Part 2.
4. Keduanya disimpan dengan `SaveAs2` pada format `wdFormatDocumentDefault` yang merupakan `.docx`.

Karena setiap keluaran berasal dari salinan biner file asli dan hanya kehilangan satu rentang konten, seluruh struktur dokumen yang tidak dihapus tetap utuh. *Styles*, *theme*, *numbering definition*, *table*, *image*, *hyperlink*, *bookmark*, *header*, *footer*, *page setup*, *section*, *font embedding*, dan *document metadata* dipertahankan oleh Word sendiri, bukan direkonstruksi oleh aplikasi ini.

### 3.4 Keterbatasan yang diketahui

Keterbatasan berikut bersifat inheren pada model dokumen Word dan tidak dapat dihilangkan oleh pendekatan mana pun.

1. **Nomor halaman otomatis** pada Part 2 kembali dimulai dari 1, kecuali dokumen memang mengatur nomor awal secara eksplisit pada *section* tersebut.
2. **Section break yang dihapus.** Bila rentang yang dihapus memuat *section break*, Word menggabungkan *section* yang berbatasan. Pada Part 2 hal ini berarti *header* dan *footer* yang berlaku adalah milik *section* yang tersisa. Perilaku ini identik dengan yang terjadi bila pengguna melakukan penghapusan yang sama secara manual di dalam Word.
3. **Referensi silang, daftar isi, dan catatan kaki** yang menunjuk ke konten di bagian lain akan menjadi rusak setelah pemisahan, karena targetnya tidak lagi berada di dalam dokumen yang sama. Perbarui *field* tersebut dengan `Ctrl` + `A` lalu `F9` setelah pemisahan.
4. **Penomoran daftar bertingkat** yang melintasi titik pisah akan melanjutkan penomoran dari definisi *numbering* yang diwarisi. Nomor awal pada Part 2 mungkin perlu disetel ulang secara manual.
5. **Pagination bergantung pada mesin.** Dokumen yang sama dapat memiliki jumlah halaman berbeda pada komputer lain apabila *font* yang digunakan tidak terpasang atau *printer driver* default berbeda. Nomor halaman yang dimasukkan pengguna harus mengacu pada tampilan Word di komputer yang sama.
6. **Dokumen terproteksi kata sandi** tidak dapat diproses karena Word akan meminta kredensial secara interaktif, dan mode otomasi menolak dialog tersebut.

---

## 4. Menjalankan dari sumber Python

```bat
py -3 -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
py -3 src\main.py
```

Menjalankan aplikasi pada sistem operasi selain Windows akan berhenti dengan pesan yang jelas, karena COM dan Microsoft Word tidak tersedia di sana.

---

## 5. Build `.exe`

```bat
build.bat
```

Skrip tersebut membuat *virtual environment* pada `build\venv`, memasang dependency, membersihkan hasil *build* sebelumnya, lalu menjalankan PyInstaller dalam mode `--onefile --windowed`. Hasil akhirnya adalah

```
dist\WordSplitter.exe
```

Executable bersifat mandiri. Pengguna akhir tidak perlu memasang Python maupun `pywin32`. Microsoft Word tetap wajib ada karena aplikasi ini menggunakan mesin pagination Word.

Beberapa produk antivirus menandai executable PyInstaller sebagai mencurigakan karena pola *self extraction*. Bila hal ini terjadi, tambahkan pengecualian untuk `dist\WordSplitter.exe` atau lakukan penandatanganan kode dengan sertifikat organisasi.

---

## 6. Cara penggunaan

1. Tekan `Select Word File` dan pilih dokumen `.docx`. Nama, lokasi, ukuran, dan status file ditampilkan.
2. Aplikasi langsung menjalankan analisis. Status berubah menjadi `Analyzing document...` lalu jumlah halaman ditampilkan.
3. Isi `Page A` dan `Page B`. Titik pisah berada di tengah antara keduanya.
4. Isi `File 1 Name` dan `File 2 Name`. Ekstensi `.docx` ditambahkan otomatis.
5. Centang `Delete original file` hanya bila file asli memang ingin dihapus. Secara *default* opsi ini tidak aktif.
6. Tekan `Process`. Bila ada nama file yang bentrok, akan muncul dialog `File already exists.` dengan pilihan `Replace`, `Choose Another Name`, dan `Cancel`.
7. Setelah selesai, kedua file hasil berada pada direktori yang sama dengan file asli.

Selama proses berjalan, antarmuka tetap responsif karena seluruh pekerjaan berat dijalankan pada *worker thread*. Tombol `Cancel` menghentikan proses pada batas tahap berikutnya dan membersihkan seluruh file sementara.

---

## 7. Jaminan integritas data

Urutan operasi dirancang agar kegagalan pada titik mana pun tidak menghasilkan kehilangan data.

```
Original
   ↓  (dibuka read-only, tidak pernah dimodifikasi)
Analyze pagination
   ↓
Salin ke temporary Part 1  →  hapus rentang  →  simpan
   ↓
Salin ke temporary Part 2  →  hapus rentang  →  simpan
   ↓
Validasi Part 1 dan Part 2 (ada, ukuran > 0, paket .docx valid, dapat dibuka ulang oleh Word)
   ↓
Commit atomik dengan os.replace ke nama final
   ↓
Hapus original (hanya bila checkbox aktif dan seluruh langkah di atas berhasil)
```

Prinsip yang diterapkan.

* File asli dibuka dalam mode `ReadOnly` dan tidak pernah menjadi sasaran operasi tulis.
* Seluruh keluaran dibentuk pada nama sementara di *volume* yang sama dengan tujuan, sehingga promosi ke nama final adalah `os.replace` yang atomik pada NTFS.
* Keberhasilan `save()` tidak dianggap sebagai bukti keberhasilan. Setiap keluaran dibuka kembali oleh Word dan dihitung ulang jumlah halamannya.
* Bila kedua keluaran ternyata masih memuat seluruh isi dokumen, proses dinyatakan gagal dan tidak ada file final yang ditulis.
* Bila proses gagal di tengah jalan, seluruh file sementara dan *lock file* Word dibersihkan.
* Bila Part 1 sudah terpromosi lalu Part 2 gagal, Part 1 dihapus kembali, kecuali bila nama tersebut sudah ditempati file milik pengguna sebelum proses dimulai. Dalam kasus itu file pengguna dibiarkan apa adanya.
* Ruang disk diperiksa sebelum proses dimulai, dengan estimasi konservatif sebesar empat kali ukuran dokumen ditambah margin.

---

## 8. Log dan diagnostik

Log ditulis ke

```
%LOCALAPPDATA%\WordSplitter\logs\wordsplitter.log
```

Bila direktori tersebut tidak dapat ditulis, log dialihkan ke `%TEMP%\WordSplitter\`. Lokasi aktif selalu ditampilkan pada bagian bawah jendela aplikasi.

File log berotasi otomatis pada ukuran 1 MB dengan tiga berkas cadangan. Isi log mencakup waktu, file masukan dan ukurannya, halaman yang dipilih, rencana pemisahan beserta *character offset*, setiap tahap yang dijalankan, seluruh kesalahan beserta *traceback* teknis, dan hasil akhir proses.

Pengguna tidak pernah melihat *traceback* mentah. Setiap kesalahan disajikan sebagai pesan yang dapat dibaca manusia, dan detail teknisnya disimpan ke log.

---

## 9. Troubleshooting

| Gejala | Penyebab dan tindakan |
|---|---|
| `Microsoft Word desktop tidak dapat dijalankan` | Word tidak terpasang, berupa versi Store atau *web*, atau lisensinya belum aktif. Buka Word secara manual satu kali untuk menyelesaikan aktivasi |
| `Microsoft Word tidak dapat membuka dokumen ini` | Dokumen terproteksi kata sandi, rusak, atau sedang dikunci proses lain. Buka dokumen di Word, simpan ulang sebagai `.docx` baru, lalu ulangi |
| `File bukan dokumen .docx yang valid` | File sebenarnya berformat `.doc` lama atau hanya diubah ekstensinya. Buka di Word lalu simpan sebagai `.docx` |
| Jumlah halaman berbeda dari tampilan Word milik pengguna | *Printer driver* default atau *font* berbeda. Pastikan dokumen dibuka pada komputer yang sama dengan yang dipakai untuk menentukan nomor halaman |
| Aplikasi tampak lambat pada dokumen sangat besar | Word merender ulang dokumen tiga kali, yaitu satu kali untuk analisis dan dua kali untuk pembentukan hasil. Ini normal. Tombol `Cancel` tetap berfungsi |
| `Ruang disk tidak mencukupi` | Sediakan ruang minimal empat kali ukuran dokumen pada *volume* tujuan |
| Proses Word tertinggal di Task Manager | Tutup aplikasi secara normal. Instance Word yang dibuat aplikasi ini selalu ditutup pada akhir setiap operasi, termasuk saat terjadi kegagalan |

---

## 10. Struktur proyek

```
WordSplitter/
├── src/
│   ├── main.py            entry point, logging, global exception hook, DPI
│   ├── gui.py             antarmuka Tkinter, worker thread, dialog konflik
│   ├── word_engine.py     lapisan COM Microsoft Word
│   ├── split_engine.py    orkestrasi, staged commit, validasi hasil, rollback
│   ├── validation.py      aturan input, nama file, aritmetika titik pisah
│   ├── utils.py           inspeksi file, integritas paket docx, temporary path
│   └── logger.py          logging berotasi
├── tests/
│   ├── test_core.py       42 pengujian logika murni
│   └── test_split_flow.py 18 pengujian orkestrasi dengan simulasi sesi Word
├── requirements.txt
├── build.bat
├── README.md
├── TEST_PLAN.md
└── dist/
    └── WordSplitter.exe   dihasilkan oleh build.bat
```

Ketergantungan antar modul bersifat satu arah, yaitu `main` → `gui` → `split_engine` → `word_engine`, dengan `validation`, `utils`, dan `logger` sebagai lapisan dasar tanpa ketergantungan pada Word maupun GUI. Susunan ini yang memungkinkan seluruh aturan validasi dan seluruh protokol orkestrasi diuji tanpa Microsoft Word.
