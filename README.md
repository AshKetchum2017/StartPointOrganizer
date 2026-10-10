# StartPointOrganizer

StartPointOrganizer adalah macro VBA untuk CorelDRAW yang membantu menyiapkan cutting line. **Organize** menetapkan arah path dan mencari Start Point pada subpath tertutup, dengan pemeriksaan tangent serta fallback handle dan midpoint. **ZOrder** mengatur urutan objek curve dalam group atau urutan subpath pada sebuah curve berdasarkan posisi baris, dari kiri atas ke kanan bawah.

Kedua perintah tersedia di `StartPointMenu` dan bekerja pada objek yang dipilih di dokumen aktif.

## Main Features

### Start Point Organizer

Pilih arah `optClockwise` atau `optCounterClockwise`, lalu jalankan `cmdOrganize`. Jika belum ada pilihan arah saat form dibuka, `optClockwise` dipilih secara bawaan.

Macro menelusuri objek pilihan, termasuk anggota group dan subgroup. Objek yang bukan curve dicoba dikonversi melalui `ConvertToCurves`. Objek yang tetap tidak menjadi curve dilewati; kegagalan konversi atau akses curve dicatat sebagai failure.

Setiap subpath diperiksa secara terpisah:

1. Subpath terbuka dilewati.
2. Arah subpath tertutup dibalik bila belum sesuai dengan pilihan Clockwise atau Counter-Clockwise.
3. Kandidat Start Point dicari menurut tiga tahap prioritas di bawah.
4. Kandidat yang dipilih dijadikan Start Point, lalu status tertutup, arah, posisi, dan kelayakannya diperiksa kembali.

### Prioritas Kandidat dan Fallback

| Prioritas | Kandidat | Syarat panjang |
| --- | --- | --- |
| 1 — Fixed Handle | Node dengan segmen curve keluar | Handle keluar minimal **0,5 mm**. |
| 2 — Relative Handle | Node dengan segmen curve keluar | Handle keluar minimal **5% dari panjang segmen keluar**. |
| 3 — Line Midpoint | Titik tengah segmen garis lurus | Panjang garis minimal **2 mm**. Midpoint yang sudah ada dan memenuhi syarat dapat digunakan kembali. |

Satu tahap diperiksa sebelum beralih ke tahap berikutnya. Kandidat Relative Handle atau Line Midpoint hanya digunakan jika tahap sebelumnya tidak menghasilkan kandidat yang memenuhi syarat. Nilai **tepat pada ambang** tetap lolos.

Dalam setiap tahap, Start Point saat ini diperiksa lebih dahulu. Jika sudah memenuhi syarat pada tahap tersebut, posisinya dipertahankan. Jika tidak, macro mencari kandidat dengan selisih sudut keluar terkecil dari **0°**. Start Point lama yang hanya memenuhi tahap fallback masih dapat diganti jika ada kandidat pada tahap yang lebih tinggi.

Jika tidak ada kandidat dan tidak ada error pembacaan, hasil dicatat sebagai **No Candidate**. Jika pencarian tidak menghasilkan kandidat tetapi ada error pembacaan, hasil dicatat terpisah sebagai pencarian tanpa hasil dengan error. Pembalikan arah yang sudah diperlukan tetap dapat terjadi.

### Pemeriksaan Tangent

Kandidat harus mempunyai tangent keluar dalam rentang **0° ± 15°**. Pada kandidat node, selisih arah tangent masuk dan keluar juga harus maksimal **15°**, untuk menghindari Start Point pada cusp dengan perubahan arah tajam.

Untuk kandidat handle:

- Segmen keluar harus berupa curve dengan handle keluar yang memenuhi ambang tahapnya.
- Segmen masuk dan keluar harus mempunyai panjang yang tidak nol.
- Jika segmen masuk berupa curve, handle masuk harus tidak nol agar tangent kedatangan dapat diperiksa. Handle masuk tidak diwajibkan memenuhi ambang **0,5 mm** atau **5%** milik handle keluar.

Untuk kandidat midpoint baru, arah diperiksa pada titik tengah garis. Titik interior garis menyediakan arah masuk dan keluar yang sama, sehingga tidak memakai sudut di endpoint garis sebagai Start Point.

Ambang panjang dihitung dalam milimeter melalui konversi dari satuan dokumen. **Organize tidak mengganti satuan dokumen.**

### Midpoint dan Pemrosesan Ulang

Node baru hanya ditambahkan saat kandidat midpoint pada segmen garis benar-benar dipilih, melalui `AddNodeAt(0.5, cdrRelativeSegmentOffset)`.

Saat dijalankan ulang, macro dapat mengenali node midpoint yang diapit dua segmen garis searah dengan panjang sama, berada di tengah kedua endpoint luar, dan mempunyai panjang gabungan minimal **2 mm**. Pemeriksaan posisi dan panjang memakai toleransi pembulatan **0,0002 mm**. Dengan demikian, midpoint yang masih memenuhi syarat dapat digunakan kembali tanpa terus menambah node.

Sebelum menambah midpoint, macro menyimpan salinan curve **setelah pembalikan arah**, jika pembalikan diperlukan. Jika penambahan, pemindahan, atau validasi Start Point midpoint gagal, curve dicoba dipulihkan ke kondisi sebelum penambahan node. Pemulihan ini tidak membatalkan pembalikan arah yang sudah terjadi sebelumnya. Proses berhenti jika pemulihan gagal.

### Ringkasan dan Penanganan Error

Ringkasan Organize mencakup jumlah shape dan subpath yang diperiksa, konversi, subpath terbuka, pembalikan arah, Start Point yang sudah benar atau berubah, penggunaan fallback **handle 5%** dan **midpoint garis**, No Candidate, error pembacaan kandidat, pemulihan midpoint, warning validasi, serta failure.

Error pembacaan satu kandidat dicatat tanpa langsung menghentikan pencarian kandidat lain. Ringkasan menyertakan error kandidat pertama dan error operasi pertama jika tersedia. Error fatal, termasuk kegagalan pemulihan midpoint atau terputusnya host, menghentikan proses.

Organize berada dalam satu `CommandGroup`. Penutupannya diperiksa, dan error operasi dibedakan dari error penutupan `CommandGroup`. Gunakan **Undo sekali** untuk membatalkan hasil operasi yang selesai dengan normal. Pemulihan otomatis yang dijelaskan di atas khusus untuk perubahan midpoint, bukan pembatalan seluruh operasi.

### ZOrder

`cmdZOrder` adalah perintah tersendiri yang saat ini ditempatkan di `StartPointMenu`. Pilihan arah path tidak diperlukan, dan ZOrder tidak otomatis menjalankan Organize.

Pengurutan memakai titik tengah bounding box. Baris dibaca dari koordinat **Y tertinggi ke terendah**; item dengan selisih Y sampai **10 mm** dari item pertama pada baris yang sama dikelompokkan menjadi satu baris. Di dalam setiap baris, urutannya dari **X terkecil ke terbesar**.

| Objek terpilih | Hasil `cmdZOrder` |
| --- | --- |
| Group | Subgroup diratakan ke group induk. Curve anggota dengan minimal **3 subpath** dipisah menjadi anggota curve satu subpath; curve dengan 1–2 subpath tetap satu objek. Anggota curve kemudian diurutkan dalam Z-Order group. |
| Curve tunggal dengan minimal 3 subpath | Urutan subpath dibangun ulang dalam curve yang sama; objek tidak dipisah menjadi beberapa shape. |
| Curve tunggal dengan 1–2 subpath | Dilewati. |
| Objek lain di tingkat pilihan atau di dalam group | Tidak ikut pengurutan curve dan dilaporkan sebagai dilewati. |

Karena ZOrder dapat meratakan subgroup dan memisahkan curve anggota group, **struktur objek hasil bisa berbeda dari pilihan awal**. Objek non-curve tidak dikonversi oleh ZOrder.

Macro memeriksa jumlah anggota dan total subpath setelah pemisahan, mencoba memulihkan selection, lalu menampilkan ringkasan. ZOrder memakai satuan milimeter sementara dan memulihkan satuan dokumen setelah selesai. Seluruh operasi berada dalam satu `CommandGroup` sehingga hasil normal dapat dibatalkan dengan satu **Undo**.

## Cara Menggunakan

### Mengatur Start Point dan Arah

1. Buka dokumen CorelDRAW dan pilih objek cutting yang ingin diproses. Group boleh dipilih.
2. Buka `StartPointMenu` dan pilih `optClockwise` atau `optCounterClockwise`.
3. Tekan `cmdOrganize`.
4. Periksa ringkasan hasil, arah subpath, posisi Start Point, serta node baru jika fallback midpoint digunakan.
5. Gunakan Undo sekali jika ingin membatalkan hasil operasi yang selesai dengan normal.

Prioritas kandidat dan ambang geometri ditetapkan dalam kode; form tidak menyediakan pengaturan manual untuk nilai **0,5 mm**, **5%**, **2 mm**, atau **15°**.

### Mengurutkan Z-Order

1. Pilih group atau curve yang mempunyai subpath.
2. Buka `StartPointMenu` dan tekan `cmdZOrder`.
3. Periksa urutan dan ringkasannya. Gunakan Undo sekali jika ingin membatalkan hasil ZOrder.

Kedua perintah memerlukan dokumen aktif dan selection yang tidak kosong. Form disembunyikan saat pesan hasil atau error ditampilkan, lalu ditutup setelah pesan tersebut selesai. `cmdCancel` menutup form tanpa memproses objek.

## Integrasi dan Source

Repository menyediakan source VBA berupa code-behind UserForm dan standard module; layout visual UserForm lengkap tidak disertakan. Siapkan UserForm `StartPointMenu` dan kontrol dengan nama sesuai kode saat memasangnya di project VBA CorelDRAW.

| Lokasi | Peran |
| --- | --- |
| `src/forms/StartPointMenu.vba` | Handler Organize/ZOrder, pilihan arah, pemrosesan shape dan subpath, pemulihan midpoint, ringkasan, dan callback penutupan form. |
| `src/modules/SPOStartPoint.bas` | Pencarian kandidat bertingkat, pemeriksaan handle/tangent, penambahan dan pengenalan midpoint, serta validasi kandidat. |
| `src/modules/SPOZOrder.bas` | Pemrosesan group, pemisahan subpath yang memenuhi syarat, pengurutan Z-Order, serta pemulihan selection dan satuan. |
| `src/modules/MRTargetBridge.bas` | Membuka UserForm dari MacroRunner melalui `OpenMacro` dan menghubungkan callback saat form ditutup. |
| `Changelog.log` | Riwayat perubahan dan catatan pengujian. |

`MRTargetBridge` menyediakan integrasi untuk alur **manual** MacroRunner tanpa memerlukan reference ke project VBA MacroRunner. Source yang tersedia belum menyediakan `ValidateBehavior` atau `RunBehavior` untuk instruksi MacroBehavior otomatis.

## Catatan

- Organize bekerja pada subpath tertutup; subpath terbuka tetap dilewati.
- Konversi objek non-curve, pemindahan Start Point, dan penambahan midpoint dapat mengubah struktur objek atau node.
- Pemeriksaan tangent dan panjang membantu memilih kandidat sesuai aturan macro; hasil tetap perlu diperiksa pada cutting line yang akan digunakan.
- Changelog **10 Oktober 2026** mencatat uji curve custom berukuran **5 × 5 sampai 150 × 150 mm** yang sesuai aturan, serta Undo yang tetap aman. Catatan tersebut mengacu pada kasus uji yang disebutkan, bukan jaminan untuk seluruh bentuk atau versi CorelDRAW.

## Lisensi dan Masukan

Project ini menggunakan **MIT License**; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
