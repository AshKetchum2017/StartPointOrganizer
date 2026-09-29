# StartPointOrganizer

StartPointOrganizer adalah macro VBA untuk CorelDRAW yang membantu menyiapkan cutting line. **Organize** menetapkan arah path dan mencari start point pada subpath tertutup. **ZOrder** mengatur urutan objek curve dalam group atau urutan subpath pada sebuah curve berdasarkan posisi baris, dari kiri atas ke kanan bawah.

Kedua perintah tersedia di `StartPointMenu` dan bekerja pada objek yang dipilih di dokumen aktif.

## Main Features

### Start Point Organizer

Pilih arah `optClockwise` atau `optCounterClockwise`, lalu jalankan `cmdOrganize`. Jika belum ada pilihan arah saat form dibuka, `optClockwise` dipilih secara bawaan.

Macro menelusuri objek pilihan, termasuk anggota group dan subgroup. Objek yang bukan curve dicoba dikonversi dengan `ConvertToCurves`; objek yang tidak berhasil menjadi curve dicatat sebagai dilewati. Setiap subpath curve diperiksa secara terpisah:

1. Subpath terbuka dilewati.
2. Arah subpath tertutup dibalik bila belum sesuai dengan pilihan Clockwise atau Counter-Clockwise.
3. Start point dinilai dari sudut control point awal pada segmen curve yang keluar dari node. Sudut yang dicari adalah `0° ± 15°`, dengan panjang control point awal minimal **0,5 mm**.
4. Jika start point saat ini memenuhi syarat, posisinya dibiarkan. Jika tidak, macro memilih node dengan selisih sudut terkecil dari `0°` di antara kandidat yang memenuhi syarat, lalu menjadikannya start point.

Segmen lurus atau control point yang lebih pendek dari ambang tersebut tidak dipakai sebagai kandidat start point. Jika tidak ada kandidat yang sesuai, start point tidak dipindahkan; pembalikan arah yang sudah diperlukan tetap dapat terjadi. Setelah perubahan, macro memeriksa lagi status tertutup, arah, dan sudut start point.

Pemrosesan memakai satuan milimeter sementara dan mencoba mengembalikan satuan dokumen setelah selesai. Ringkasan hasil menampilkan jumlah shape, subpath, konversi, pembalikan arah, start point yang berubah atau sudah benar, subpath terbuka, warning validasi, dan failure. Jika ada kegagalan pada satu shape atau subpath, ringkasan juga memuat error pertama yang tercatat.

### ZOrder

`cmdZOrder` adalah perintah tersendiri yang saat ini ditempatkan di `StartPointMenu`. Menjalankannya **tidak** otomatis menjalankan `cmdOrganize`, dan pilihan arah path tidak diperlukan. Pilih group atau curve terlebih dahulu.

Pengurutan memakai titik tengah bounding box. Baris dibaca dari koordinat **Y tertinggi ke terendah**; item dengan selisih Y sampai **10 mm** dari item pertama pada baris yang sama dikelompokkan menjadi satu baris. Di dalam setiap baris, urutannya dari **X terkecil ke terbesar**.

| Objek terpilih | Hasil `cmdZOrder` |
| --- | --- |
| Group | Subgroup diratakan ke group induk. Curve anggota dengan minimal **3 subpath** dipisah menjadi anggota curve satu subpath; curve dengan 1–2 subpath tetap satu objek. Anggota curve kemudian diurutkan dalam Z-order group. |
| Curve tunggal dengan minimal 3 subpath | Urutan subpath dibangun ulang dalam curve yang sama; objek tidak dipisah menjadi beberapa shape. |
| Curve tunggal dengan 1–2 subpath | Dilewati. |
| Objek lain di tingkat pilihan atau di dalam group | Tidak ikut pengurutan curve dan dilaporkan sebagai dilewati. |

Karena ZOrder dapat meratakan subgroup dan memisahkan curve anggota group, **struktur objek hasil bisa berbeda dari pilihan awal**. Macro memeriksa jumlah anggota dan total subpath setelah pemisahan, mencoba memulihkan selection, lalu menampilkan ringkasan. Seluruh operasi ZOrder berada dalam satu `CommandGroup` sehingga hasilnya dapat dibatalkan dengan satu **Undo**.

## Cara menggunakan

### Mengatur start point dan arah

1. Buka dokumen CorelDRAW dan pilih objek cutting yang ingin diproses. Group boleh dipilih.
2. Buka `StartPointMenu` dan pilih `optClockwise` atau `optCounterClockwise`.
3. Tekan `cmdOrganize`, lalu periksa ringkasan hasil dan posisi start point pada curve.

### Mengurutkan Z-order

1. Pilih group atau curve yang mempunyai subpath.
2. Buka `StartPointMenu` dan tekan `cmdZOrder`.
3. Periksa urutan visual dan ringkasannya. Gunakan **Undo** sekali jika ingin membatalkan hasil ZOrder.

Kedua perintah memerlukan dokumen aktif dan selection yang tidak kosong. Form ditutup setelah pesan hasil atau error ditampilkan. `cmdCancel` menutup form tanpa memproses objek.

## Integrasi dan source

Repository menyediakan source VBA berupa code-behind UserForm dan standard module; layout visual UserForm lengkap tidak disertakan. Siapkan form dan kontrol dengan nama sesuai kode saat memasangnya di project VBA CorelDRAW.

| Lokasi | Peran |
| --- | --- |
| `src/forms/StartPointMenu.vba` | Handler `cmdOrganize`, `cmdZOrder`, pilihan arah, ringkasan, dan callback penutupan form. |
| `src/modules/SPOZOrder.bas` | Pemrosesan group, pemisahan subpath yang memenuhi syarat, dan pengurutan Z-order. |
| `src/modules/MRTargetBridge.bas` | Membuka UserForm dari MacroRunner dan menghubungkan callback saat form ditutup. |

`MRTargetBridge` di repository ini menyediakan `OpenMacro` untuk alur **manual** MacroRunner. Kode yang tersedia belum menyediakan `ValidateBehavior` atau `RunBehavior` untuk instruksi MacroBehavior otomatis.

## Lisensi dan masukan

Project ini menggunakan lisensi MIT; lihat [LICENSE](LICENSE).

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
