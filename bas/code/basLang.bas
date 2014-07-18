Attribute VB_Name = "basLang"
' ######################
' Disini neh translasi bahasa

' code
' a_xxxx() ' untuk tulisan yang nempel pada objek tombol
' b_xxxx() ' untuk tulisan yang nempel pada objek tab
' c_xxxx() ' untuk tulisan yang nempel pada objek frame
' d_xxxx() ' untuk tulisan yang nempel pada objek label
' e_xxxx() ' untuk tulisan yang nempel pada objek listview luar
' f_xxxx() ' untuk tulisan yang nempel pada objek listview dalam
' g_xxxx() ' untuk tulisan yang nempel pada objek menu
' h_xxxx() ' untuk tulisan yang nempel pada objek checkbox
' i_xxxx() ' untuk tulisan yang nempel pada objek msgbox dan baloon
' j_xxxx() ' untuk tulisan yang tertinggal

Public a_bahasa(30)   As String
Public b_bahasa(30)   As String
Public c_bahasa(10)   As String
Public d_bahasa(40)   As String
Public e_bahasa(20)   As String
Public f_bahasa(22)   As String
Public g_bahasa(15)   As String
Public h_bahasa(10)   As String
Public i_bahasa(30)   As String
Public j_bahasa(70)   As String



Public Sub InitLanguange(LangName As String)
Dim LangPath As String
LangPath = GetFilePath(App_FullPathW(False)) & "\lang"

Select Case LangName
    Case "Default@1": Call InitInternalLang ' ENG
    Case "Default@2": Call InitInternalLang2 ' IND
    Case Else: GoTo LBL_CARI2
End Select

GoTo LBL_AKHIR

LBL_CARI2:
If ValidFile(LangPath & "\" & LangName) = True Then ' jika ada
   If ReadExternalLang(LangPath & "\" & LangName) = False Then
      GoTo LBL_FALSE
   End If
Else
  GoTo LBL_FALSE
End If

GoTo LBL_AKHIR ' klo gagal init extrenal (bisa saja bahasa extrenal tidak valid)

LBL_FALSE:
    Call InitInternalLang ' ENG ' balik default jika gagal
    MsgBox j_bahasa(21), vbExclamation

LBL_AKHIR:
   Call WriteLangToInterface ' tulis ke interface
End Sub

Private Sub InitInternalLang2()
    a_bahasa(0) = "Pindai":                 a_bahasa(1) = "Konfigurasi"
    a_bahasa(2) = "Alat":                   a_bahasa(3) = "Perbaharui"
    a_bahasa(4) = "Tentang"
    
    a_bahasa(5) = "Mulai Scan":             a_bahasa(6) = "Lewati Buffer.."
    a_bahasa(7) = "Keluar Pindai":          a_bahasa(8) = "Fix Tercentang"
    a_bahasa(9) = "Fix Semua Objek":        a_bahasa(10) = "Properties"
    a_bahasa(11) = "Jelajah":               a_bahasa(12) = "Simpan Konfigurasi"
    a_bahasa(13) = "Terapkan Bahasa"
    
    a_bahasa(14) = "Cabut Semua":           a_bahasa(15) = "Tambah Path"
    a_bahasa(16) = "Tambah File":           a_bahasa(17) = "Eksekusi Plugin"
    a_bahasa(18) = "Klik Untuk Info Lanjut"
    
    a_bahasa(19) = "Tambah":                a_bahasa(20) = "Batal"
    
    a_bahasa(21) = "Bersihkan Penjara":     a_bahasa(22) = "Bunuh Tahanan"
    a_bahasa(23) = "Lepaskan..":            a_bahasa(24) = "Lepas Ke.."
    '-----
    a_bahasa(25) = "Fix Semua":             a_bahasa(26) = "Fix Terpilih"
    a_bahasa(27) = "Abaikan":               a_bahasa(28) = "Tutup"
    
    
    '#######
    b_bahasa(0) = "Path Pindai":            b_bahasa(1) = "Malware"
    b_bahasa(2) = "Registri":               b_bahasa(3) = "Tersembunyi"
    b_bahasa(4) = "Informasi"
    
    b_bahasa(5) = "Aplikasi":               b_bahasa(6) = "Bahasa"
    b_bahasa(7) = "Pengecualian RTP":       b_bahasa(8) = "Pengecualian File"
    b_bahasa(9) = "Pengecualian Registri"
    b_bahasa(10) = "Plugin"
    
    b_bahasa(11) = "Proses Manager":        b_bahasa(12) = "Penanda Malware Sementara"
    b_bahasa(13) = "Pengontrol Penjara"

    b_bahasa(14) = "Tentang CMC":           b_bahasa(15) = "Informasi CMC"
   
       '########
    c_bahasa(0) = "Pilihan Pemindaian"
    c_bahasa(1) = "Konfigurasi Aplikasi"
    c_bahasa(2) = "Daftar Proses"
    c_bahasa(3) = "Daftar Module"
    c_bahasa(4) = "Malware Sementara"
    c_bahasa(5) = "Daftar Temporary"
    c_bahasa(6) = "Tahanan"
    c_bahasa(7) = "Daftar Malware Dapat Dideteksi"
    c_bahasa(8) = "Informasi Software"
    c_bahasa(9) = "Detektor Virus Internal"
    
    '########
    d_bahasa(0) = "Status     :"
    d_bahasa(1) = "Diproses  :"
    d_bahasa(2) = "Waktu"
    d_bahasa(3) = "Ditemukan"
    d_bahasa(4) = "Diperiksa"
    d_bahasa(5) = "Dilalui"
    d_bahasa(6) = "Malware"
    d_bahasa(7) = "Registri"
    d_bahasa(8) = "Informasi"
    d_bahasa(9) = "[Siap]"
    d_bahasa(10) = "Pindai Registri !"
    d_bahasa(11) = "Pindai Servis !"
    d_bahasa(12) = "Pindai Proses !"
    d_bahasa(13) = "Pindai Startup !"
    d_bahasa(14) = "Pindai Root !"
    d_bahasa(15) = "Sedang Pindai File !"
    d_bahasa(16) = "Dikeluarkan"
    d_bahasa(17) = "Diselesaikan"
    
    d_bahasa(18) = "Bahasa Terpilih"
    d_bahasa(19) = "ID Bahasa"
    d_bahasa(20) = "Pengarang Bahasa"
    d_bahasa(21) = "Jangan beri saya peringatan tentang ancaman dalam path dibawah ini !"
    d_bahasa(22) = "Saya yakin ini adalah file-file normal, jangan tangkap sebagai Malware file-file dibawah !"
    d_bahasa(23) = "Saya yakin ini adalah nilai normal, Jangan tangkap sebagai nilai yang jelek daftar dibawah !"
    d_bahasa(24) = "Plugin Tersedia"
    d_bahasa(25) = "Bahasa Tersedia"
    d_bahasa(26) = "Plugin Terpilih"
    d_bahasa(27) = "Pengarang Plugin"
    d_bahasa(28) = "Deskripsi Plugin"
    d_bahasa(29) = "Alamat Malware"
    d_bahasa(30) = "Nama Malware"
    d_bahasa(31) = "Versi Engine"
    d_bahasa(32) = "Nomor Bangun"
    d_bahasa(33) = "Tgl Bangun"
    d_bahasa(34) = "Database Reg"
    d_bahasa(35) = "Penanda Worm"
    d_bahasa(36) = "Penanda Virus"
    d_bahasa(37) = "Mesin"
    
    d_bahasa(38) = "objek"
    
    '########
    e_bahasa(0) = "Nama Malware"
    e_bahasa(1) = "Alamat Objek"
    e_bahasa(2) = "Ukuran [B]"
    e_bahasa(3) = "Informasi"
    
    e_bahasa(4) = "Nama Value"
    e_bahasa(5) = "Alamat Kunci"
    
    e_bahasa(6) = "Nama Objek"
    e_bahasa(7) = "Nama File"
    
    e_bahasa(8) = "Nama Proses"
    e_bahasa(9) = "Startup"
    e_bahasa(10) = "AyahPID"
    e_bahasa(11) = "Alamat TerUpdate"
    e_bahasa(12) = "Tersembunyi"
    e_bahasa(13) = "Di Debug"
    e_bahasa(14) = "Dikunci"

    
    e_bahasa(15) = "Nama Virus"
    e_bahasa(16) = "Alamat Asli"
    e_bahasa(17) = "Dalam Penjara"

    '########
    f_bahasa(0) = "Tesembunyi"
    f_bahasa(1) = "Disangka dengan"
    f_bahasa(2) = "File tersangka"
    f_bahasa(3) = "File PE buruk"
    f_bahasa(4) = "File PE kotor"
    f_bahasa(5) = "Mengandung terlalu banyak bite tambahan, berpotensi program dropper atau installer"
    f_bahasa(6) = "File terinfeksi"
    f_bahasa(7) = "File Virus"
    f_bahasa(8) = "Startup Malware"
    f_bahasa(9) = "Value tak  berguna, seharusnya di hapus"
    f_bahasa(10) = "Value Dihapus"
    f_bahasa(11) = "String Dibenahi"
    f_bahasa(12) = "DWORD Dibenahi"
    f_bahasa(13) = "File Dinormalkan"
    f_bahasa(14) = "Folder Dinormalkan"
    f_bahasa(15) = "Dikembalikan"
    f_bahasa(16) = "Dikirim ke penjara"
    f_bahasa(17) = "Dikirim ke penjara, tapi gagal dicabut sumbernya"
    f_bahasa(18) = "Gagal dipenjara dan dicabut sumbernya"
    f_bahasa(19) = "File milik Sistem"
    f_bahasa(20) = "File dibutuhkan oleh Sistem agar berjalan normal"
    f_bahasa(21) = "Mengandung banyak bite tambahan, bisa saja data Anda, "
    
    '########
    g_bahasa(0) = "Sembunyikan Pemindai"
    g_bahasa(1) = "CMC Pemindai"
    g_bahasa(2) = "Aktifkan Pelindung"
    g_bahasa(3) = "Jalan pada Startup"
    g_bahasa(4) = "&Keluar"
    
    g_bahasa(5) = "Fix Terpilih"
    g_bahasa(6) = "Fix Tercawang"
    g_bahasa(7) = "Fix Semua Objek"
    g_bahasa(8) = "Tambahkan Pengecualian"
    g_bahasa(9) = "Jelajah Objek"
    g_bahasa(10) = "Properties"
    
    g_bahasa(11) = "Segarkan Proses"
    g_bahasa(12) = "Bunuh Proses"
    g_bahasa(13) = "Jalan Ulang Proses"
    g_bahasa(14) = "Pause Proses"
    g_bahasa(15) = "Resume Proses"

    '########
    h_bahasa(0) = "Perbolehkan saring file (melewati file dengan ekstensi kusus)"
    h_bahasa(1) = "Perbolehkan menggunakan Heuristic untuk mencurigai virus"
    h_bahasa(2) = "Perbolehkan deteksi nilai registry yang tak berguna (Hanya XP)"
    h_bahasa(3) = "Perbolehkan deteksi objek yang tersembunyi (file dan folder)"
    h_bahasa(4) = "Perbolehkan memberikan informasi selama pemindaian berjalan"
    h_bahasa(5) = "Perbolehkan unpack file arsip (zip, rar, gz, tgz)"
    
    h_bahasa(6) = "Perbolehkan jalan pada Startup"
    h_bahasa(7) = "Aktifkan Proteksi CMC"
    h_bahasa(8) = "Cek Perbaharuan Online Automatis"
    h_bahasa(9) = "Scan FD yang masuk Automatis"
    h_bahasa(10) = "Tempatkan aplikasi di atas"

    '########
    i_bahasa(0) = "Semua tahanan dibunuh !"
    i_bahasa(1) = "File sudah ada. apakah anda ingin menumpuknya ?"
    i_bahasa(2) = "Tahanan dilepaskan ke"
    i_bahasa(3) = "Path asli tidak tersedia - gunakan path tertentu to untuk melepaskan tahanan !"
    i_bahasa(4) = "Berhasil meng-unload module terpilih"
    i_bahasa(5) = "Gagal meng-unload module terpilih"
    i_bahasa(6) = "Mungkin bekerja dengan baik setelah aplikasi di jalankan ulang !"
    i_bahasa(7) = "Pilih file yang akan di tandai sebagai malware sementara !"
    i_bahasa(8) = "Mohon mengeluarkan proses pindai terlebih dulu !"
    i_bahasa(9) = "Apakah anda yakin untuk meng-unload module terpilih ?"
    i_bahasa(10) = "Proses dengan PID"
    i_bahasa(11) = "Berhasil dimatikan dengan baik !"
    i_bahasa(12) = "tidak dapat dimatikan !"
    i_bahasa(13) = "sukses dijalankan ulang !"
    i_bahasa(14) = "tidak dapat dijalankan ulang !"
    i_bahasa(15) = "Jalankan ulang program untuk menerapkan semua pengaturan !"
    i_bahasa(16) = "Berhasil ditambahkan sebagai contoh malware secara sementara !"
    i_bahasa(17) = "Nama baru malware anda"
    i_bahasa(18) = "Pindai komputer untuk melihat hasilnya"
    i_bahasa(19) = "Apakah anda yakin to membersihkan semua tahanandari penjara CMC ?"
    i_bahasa(20) = "Gagal ditambahkan sebagai malware baru anda secara sementara"
    i_bahasa(21) = "Apakah anda yakin untuk membunuh tahanan-tahanan terpilih"
    i_bahasa(22) = "Apakah anda yakin untuk melepaskan tahanan terpilih"
    i_bahasa(23) = "CMC gagal mendapatkan daftar file milik sistem, ini dapat membuat CMC mengahapus file virus walaupun dibutuhkan oleh sistem"
    i_bahasa(24) = "Proteksi CMC berstatus dinyalakan, Sistem anda sekarang dilindungi oleh C.M.C+"
    i_bahasa(25) = "Proteksi CMC berstatus dimatikan, C.M.C+ istirahat dalam melindungi Sistem anda"
    i_bahasa(26) = "Informasi"
    i_bahasa(27) = "Perhatian"
    
    '########
    j_bahasa(0) = "Pindai Module !"
    j_bahasa(1) = "Sevice gagal di musnahkan !"
    j_bahasa(2) = "Servis Dimusnahkan"
    j_bahasa(3) = "Servis"
    j_bahasa(4) = "Ditemukan di Reg-Startup"
    j_bahasa(5) = "Startup Virus"
    j_bahasa(6) = "Ditemukan di Startup-Explorer"
    j_bahasa(7) = "di Memori [Dimatikan+Dikunci]"
    j_bahasa(8) = "[Dikeluarkan dari Proses]"
    j_bahasa(9) = "Virus-Ku"
    j_bahasa(10) = "Daerah Sistem"
    j_bahasa(11) = "Proses dan Servis"
    j_bahasa(12) = "nilai"
    j_bahasa(13) = "item tercentang"
    j_bahasa(14) = "item terpilih"
    j_bahasa(15) = "semua item"
    j_bahasa(16) = "Apakah anda yakin untuk membenahi"
    j_bahasa(17) = "yang akan dibenahi"
    j_bahasa(18) = "yang sudah dibenahi"
    j_bahasa(19) = "Sebagian item dari malware yang terdeteksi belum dibenahi."
    j_bahasa(20) = "Anda akan kehilangan informasi jika melanjutkan proses pemindaian, Apakah Anda yakin ?"
    j_bahasa(21) = "Gagal membaca bahasa external !"
    j_bahasa(22) = "Tidak dapat mengenumerisasi module dari proses terpilih"
    j_bahasa(23) = "Cabut Terpilih"
    j_bahasa(24) = "Gagal mendapatkan informasi Pembaharuan !"
    j_bahasa(25) = "Tidak ada pembaharuan baru tersedia untuk versi CMC Anda"
    j_bahasa(26) = "Sedang Mendapatkan Info Pembaharuan ..."
    j_bahasa(27) = "Perbaharui Sekarang"
    j_bahasa(28) = "Sebagian komponen penting gagal dibaca"
    j_bahasa(29) = "Memperbaharui"
    j_bahasa(30) = "Mengunduh"
    j_bahasa(31) = "Selesai.."
    j_bahasa(32) = "Perbaharuan Komponen Dibatalkan.."
    j_bahasa(33) = "Sedang Cek Pembaharuan.."
    j_bahasa(34) = "Batalkan Pembaharuan"
    j_bahasa(35) = "[Path Kusus]"
    j_bahasa(36) = "Bahasa Dipakai"
    j_bahasa(37) = "Sedang Menyangga"
    j_bahasa(38) = "file"
    j_bahasa(39) = "folder"
    j_bahasa(40) = "Anda tidak memakai Sistem Operasi WinXP, tolong matikan fitur"
    j_bahasa(41) = "LAPORAN-PINDAI"
    j_bahasa(42) = "Status Pindai"
    j_bahasa(43) = "File Ditemukan"
    j_bahasa(44) = "File Dipindai"
    j_bahasa(45) = "File Tidak Dipindai"
    j_bahasa(46) = "Virus Ditemukan"
    j_bahasa(47) = "Value Dipindai"
    j_bahasa(48) = "Value Buruk"
    j_bahasa(49) = "Waktu Berakhir"
    j_bahasa(50) = "Pasang Konteks Menu -Pindai Dengan"
    j_bahasa(51) = "Pindai dengan"
    j_bahasa(52) = "Email Pengarang"
    j_bahasa(53) = "Web Pengarang"
    j_bahasa(54) = "Kode Verifikasi"
    j_bahasa(55) = "Tidak ada plugin tersedia - dapatkan plugin CMC di"
    j_bahasa(56) = "Apakah anda yakin untuk menjalankan plugin terpilih"
    j_bahasa(57) = "Jalankan sebagai threat baru"
    j_bahasa(58) = "Plugin tidak dapat dijalankan"
    j_bahasa(59) = "menemukan virus saat pertama berjalan"
    j_bahasa(60) = "Removable drive terdeteksi"
    j_bahasa(61) = "mendeteksi removable drive dimasukan, apakah anda ingin memindai dengan"
    j_bahasa(62) = "Memindai removable drive yang masuk"


End Sub

Private Sub InitInternalLang()
    a_bahasa(0) = "Scan":                   a_bahasa(1) = "Configuration"
    a_bahasa(2) = "Tool":                   a_bahasa(3) = "Update"
    a_bahasa(4) = "About"
    
    a_bahasa(5) = "Start Scan":             a_bahasa(6) = "Skip Buffer.."
    a_bahasa(7) = "Abort Scan":             a_bahasa(8) = "Fix Checked"
    a_bahasa(9) = "Fix All Object":         a_bahasa(10) = "Properties"
    a_bahasa(11) = "Explore":               a_bahasa(12) = "Save Configuration"
    a_bahasa(13) = "Apply Language"
    
    a_bahasa(14) = "Remove All":            a_bahasa(15) = "Add Path"
    a_bahasa(16) = "Add File":              a_bahasa(17) = "Execute Plugin"
    a_bahasa(18) = "Click For More Information"
    
    a_bahasa(19) = "Add":                   a_bahasa(20) = "Cancel"
    
    a_bahasa(21) = "Clear Jail":            a_bahasa(22) = "Kill Prisoner"
    a_bahasa(23) = "Release..":             a_bahasa(24) = "Release To.."
    '-----
    a_bahasa(25) = "FIX ALL":               a_bahasa(26) = "FIX Selected"
    a_bahasa(27) = "Ignore":                a_bahasa(28) = "Close"
    
    '#######
    b_bahasa(0) = "Path Scan":              b_bahasa(1) = "Malware"
    b_bahasa(2) = "Registry":               b_bahasa(3) = "Hidden"
    b_bahasa(4) = "Information"
    
    b_bahasa(5) = "Application":            b_bahasa(6) = "Language"
    b_bahasa(7) = "RTP Exception(s)":       b_bahasa(8) = "File Exception(s)"
    b_bahasa(9) = "Registry Exception(s)"
    b_bahasa(10) = "Plugin(s)"
    
    b_bahasa(11) = "Prosess Manager":       b_bahasa(12) = "Temporary Malware Signer"
    b_bahasa(13) = "Jail Controller"

    b_bahasa(14) = "About CMC":             b_bahasa(15) = "CMC Information"
   
    '########
    c_bahasa(0) = "Scan Option"
    c_bahasa(1) = "Application Configuration"
    c_bahasa(2) = "Process List"
    c_bahasa(3) = "Module List"
    c_bahasa(4) = "Malware Temporary"
    c_bahasa(5) = "List of Temporary"
    c_bahasa(6) = "Prisoner"
    c_bahasa(7) = "List of Detected Malware"
    c_bahasa(8) = "Software Information"
    c_bahasa(9) = "Internal Virus Detector"
    
    '########
    d_bahasa(0) = "Status       :"
    d_bahasa(1) = "Processed :"
    d_bahasa(2) = "Time"
    d_bahasa(3) = "Founded"
    d_bahasa(4) = "Checked"
    d_bahasa(5) = "ByPassed"
    d_bahasa(6) = "Malware"
    d_bahasa(7) = "Registry"
    d_bahasa(8) = "Information"
    d_bahasa(9) = "[Ready]"
    d_bahasa(10) = "Scan Registry !"
    d_bahasa(11) = "Scan Service !"
    d_bahasa(12) = "Scan Process !"
    d_bahasa(13) = "Scan Startup !"
    d_bahasa(14) = "Scan Root !"
    d_bahasa(15) = "Scanning File !"
    d_bahasa(16) = "Aborted"
    d_bahasa(17) = "Finished"
    
    d_bahasa(18) = "Language Selected"
    d_bahasa(19) = "Language ID"
    d_bahasa(20) = "Language Author"
    d_bahasa(21) = "Dont Give me Warning about Threat in this Path"
    d_bahasa(22) = "I'am sure this is normal file, dont catch as a malware file(s) below"
    d_bahasa(23) = "I'am sure this is normal value. Don't catch as a bad value, value(s) below"
    d_bahasa(24) = "Avalaible Plugin(s)"
    d_bahasa(25) = "Avalaible Language"
    d_bahasa(26) = "Plugin Selected"
    d_bahasa(27) = "Plugin Author"
    d_bahasa(28) = "Plugin Description"
    d_bahasa(29) = "Malware Path"
    d_bahasa(30) = "Malware Name"
    d_bahasa(31) = "Engine Version"
    d_bahasa(32) = "Build Number"
    d_bahasa(33) = "Build Date"
    d_bahasa(34) = "Reg Database"
    d_bahasa(35) = "Worm Signature"
    d_bahasa(36) = "Virus Signature"
    d_bahasa(37) = "Machine"
    
    d_bahasa(38) = "object(s)"
    '########
    e_bahasa(0) = "Malware Name"
    e_bahasa(1) = "Object Path"
    e_bahasa(2) = "Size [B]"
    e_bahasa(3) = "Information"
    
    e_bahasa(4) = "Value Name"
    e_bahasa(5) = "Key Path"
    
    e_bahasa(6) = "Object Name"
    e_bahasa(7) = "File Name"
    
    e_bahasa(8) = "Process Name"
    e_bahasa(9) = "Startup"
    e_bahasa(10) = "ParentPID"
    e_bahasa(11) = "Update Path"
    e_bahasa(12) = "Hidden"
    e_bahasa(13) = "In Debug"
    e_bahasa(14) = "Locked"

    
    e_bahasa(15) = "Virus Name"
    e_bahasa(16) = "Original Path"
    e_bahasa(17) = "In jail"
    
    '########
    f_bahasa(0) = "Hidden"
    f_bahasa(1) = "Suspected With"
    f_bahasa(2) = "Suspected File"
    f_bahasa(3) = "Bad PE File"
    f_bahasa(4) = "Dirty PE File"
    f_bahasa(5) = "Contain too much additonal bytes - Potensial Dropper/Installer (please send to us if you also suspect it)"
    f_bahasa(6) = "Infected file"
    f_bahasa(7) = "Virus file"
    f_bahasa(8) = "Malware Startup"
    f_bahasa(9) = "Useless Value, Should be deleted"
    f_bahasa(10) = "Value Deleted"
    f_bahasa(11) = "String Fixed"
    f_bahasa(12) = "DWORD Fixed"
    f_bahasa(13) = "File Normalized"
    f_bahasa(14) = "Folder Normalized"
    f_bahasa(15) = "Restored"
    f_bahasa(16) = "Sent to jail !"
    f_bahasa(17) = "Sent to jail but fail remove source !!"
    f_bahasa(18) = "Fail sent to jail and remove source !!"
    f_bahasa(19) = "File System"
    f_bahasa(20) = "File is needed by system to run normally !"
    f_bahasa(21) = "Contain too much additional bytes (maybe your data), "
    
     '########
    g_bahasa(0) = "Hide Scanner"
    g_bahasa(1) = "CMC Scanner"
    g_bahasa(2) = "Enable Protection"
    g_bahasa(3) = "Run On Startup"
    g_bahasa(4) = "&Exit"
    
    g_bahasa(5) = "Fix Selected"
    g_bahasa(6) = "Fix Checked"
    g_bahasa(7) = "Fix All Object"
    g_bahasa(8) = "Add Exception"
    g_bahasa(9) = "Explore Object"
    g_bahasa(10) = "Properties"
    
    g_bahasa(11) = "Refresh Process"
    g_bahasa(12) = "Kill Process"
    g_bahasa(13) = "Restart Process"
    g_bahasa(14) = "Pause Process"
    g_bahasa(15) = "Resume Process"

    '######## - CB
    h_bahasa(0) = "Enable filter file (by pass file with certain extensions)"
    h_bahasa(1) = "Enable use Heuristic to suspect malware"
    h_bahasa(2) = "Enable detect useless registry value (XP only)"
    h_bahasa(3) = "Enable detect hidden object (file and folder)"
    h_bahasa(4) = "Enable give strange  information while scanning"
    h_bahasa(5) = "Enable unpack archive (zip, rar, gz, tgz)"
    
    h_bahasa(6) = "Enable Run on Startup"
    h_bahasa(7) = "Enable CMC Protection"
    h_bahasa(8) = "Auto Check Online Update"
    h_bahasa(9) = "Auto Scan Flashdisk inserted"
    h_bahasa(10) = "Place Application on Top"
    
    '########
    i_bahasa(0) = "All prisoner killed !"
    i_bahasa(1) = "File is already exist. Do you want to over write?"
    i_bahasa(2) = "Prisoner released to"
    i_bahasa(3) = "Original path is not avalaible  - use custom path to release prisoner !"
    i_bahasa(4) = "Success unload selected module"
    i_bahasa(5) = "Fail to unload selected module"
    i_bahasa(6) = "Maybe work well after application restarted !"
    i_bahasa(7) = "Select a file to be signed as your malware !"
    i_bahasa(8) = "Please terminate scanning process first !"
    i_bahasa(9) = "Are you sure to unload selected module?"
    i_bahasa(10) = "Process with PID"
    i_bahasa(11) = "was terminated succesfully !"
    i_bahasa(12) = "cannot be terminated !"
    i_bahasa(13) = "was restarted succesfully !"
    i_bahasa(14) = "cannot be restarted !"
    i_bahasa(15) = "Restart Application for apply all change"
    i_bahasa(16) = "success added as new temporary malware sample !"
    i_bahasa(17) = "Your new malware name"
    i_bahasa(18) = "Scan computer to view the result"
    i_bahasa(19) = "Are you sure to clear all prisoner in C.M.C jail ?"
    i_bahasa(20) = "Fail added as new temporary malware"
    i_bahasa(21) = "Are you sure to kill selected prisoners"
    i_bahasa(22) = "Are you sure to release selected prisoner"
    i_bahasa(23) = "CMC fail to get file system list, it can make C.M.C delete virus file although needed by your system"
    i_bahasa(24) = "CMC Protector is turn ON, your system are protected by C.M.C+ now"
    i_bahasa(25) = "CMC Protector is turn OFF, C.M.C+ is rest to protect your system"
    i_bahasa(26) = "Information"
    i_bahasa(27) = "Caution"
    
    
    '########
    j_bahasa(0) = "Scan Module !"
    j_bahasa(1) = "Service Fail Destroyed"
    j_bahasa(2) = "Service Destroyed"
    j_bahasa(3) = "Service"
    j_bahasa(4) = "Found in Reg-Startup"
    j_bahasa(5) = "Virus Startup"
    j_bahasa(6) = "Found in Explorer-Startup"
    j_bahasa(7) = "in Memory [Terminated+Locked]"
    j_bahasa(8) = "[Unload From Process]"
    j_bahasa(9) = "My-Virus"
    j_bahasa(10) = "System Areas"
    j_bahasa(11) = "Process + Service"
    j_bahasa(12) = "value(s)"
    j_bahasa(13) = "checked item"
    j_bahasa(14) = "selected item"
    j_bahasa(15) = "all item"
    j_bahasa(16) = "Are you sure to fixed"
    j_bahasa(17) = "will be fixed"
    j_bahasa(18) = "has been fixed"
    j_bahasa(19) = "Some items of detected malware has not been fixed."
    j_bahasa(20) = "You will lost information if you continue scanning process, Are you sure ?"
    j_bahasa(21) = "Fail read external language !"
    j_bahasa(22) = "Cannot enumerate modules from selected process"
    j_bahasa(23) = "Remove Selected"
    j_bahasa(24) = "Fail to get Update information !"
    j_bahasa(25) = "No New Update avalaible for your CMC version"
    j_bahasa(26) = "Getting Update Info ...."
    j_bahasa(27) = "Update Now"
    j_bahasa(28) = "Some of important component is fail to read"
    j_bahasa(29) = "Updating"
    j_bahasa(30) = "Downloading"
    j_bahasa(31) = "Done.."
    j_bahasa(32) = "Update Component Canceled.."
    j_bahasa(33) = "Checking Update.."
    j_bahasa(34) = "Cancel Update"
    j_bahasa(35) = "[Special Paths]"
    j_bahasa(36) = "Language Used"
    j_bahasa(37) = "Buffering"
    j_bahasa(38) = "file(s)"
    j_bahasa(39) = "folder(s)"
    j_bahasa(40) = "You are not use WinXP OS, please turn off"
    j_bahasa(41) = "SCAN-REPORT"
    j_bahasa(42) = "Scan Status"
    j_bahasa(43) = "File Found"
    j_bahasa(44) = "File Scanned"
    j_bahasa(45) = "File Not Scanned"
    j_bahasa(46) = "Virus Found"
    j_bahasa(47) = "Value Scanned"
    j_bahasa(48) = "Bad Value"
    j_bahasa(49) = "End Time"
    j_bahasa(50) = "Enable Context Menu -Scan With"
    j_bahasa(51) = "Scan With"
    j_bahasa(52) = "Author Email"
    j_bahasa(53) = "Author Site"
    j_bahasa(54) = "Verification Code"
    j_bahasa(55) = "No Plugin Avalaible - Get CMC plugin at"
    j_bahasa(56) = "Are you sure to execute selected plugin"
    j_bahasa(57) = "Run as new Thread"
    j_bahasa(58) = "Plugin can't be executed"
    j_bahasa(59) = "found virus in the first loading"
    j_bahasa(60) = "Removable drive detected"
    j_bahasa(61) = "detect new removable drive inserted, would you like to scan with"
    j_bahasa(62) = "Scanning removable drive inserted"

 End Sub


Private Sub WriteLangToInterface()
On Error Resume Next
With frmMain
     
     .bMenu(0).Caption = a_bahasa(0)
     .bMenu(1).Caption = a_bahasa(1)
     .bMenu(2).Caption = a_bahasa(2)
     .bMenu(3).Caption = a_bahasa(3)
     .bMenu(4).Caption = a_bahasa(4)
     .bMenu(5).Caption = b_bahasa(10)
     
     .cmdStartScan.Caption = a_bahasa(5)
     
     .cmdFixMalware.Caption = a_bahasa(8)
     .cmdFixMalwareAll.Caption = a_bahasa(9)
     .cmdFixReg.Caption = a_bahasa(8)
     .cmdFixRegAll.Caption = a_bahasa(9)
     .cmdFixHidden.Caption = a_bahasa(8)
     .cmdFixHiddenAll.Caption = a_bahasa(9)
     .cmdProperties.Caption = a_bahasa(10)
     .cmdExplore.Caption = a_bahasa(11)
     .cmdSave.Caption = a_bahasa(12)
     .cmdApplyLang.Caption = a_bahasa(13)
     .cmdRemovePath.Caption = a_bahasa(14)
     .cmdRemovePath1.Caption = j_bahasa(23)
     .cmdRemExcFile.Caption = a_bahasa(14)
     .cmdRemExcFile1.Caption = j_bahasa(23)
     .cmdRemExcReg.Caption = a_bahasa(14)
     .cmdRemExcReg1.Caption = j_bahasa(23)

     .cmdAddExcFolder.Caption = a_bahasa(15)
     .cmdAddExcFile.Caption = a_bahasa(16)
     .cmdExecutePlug.Caption = a_bahasa(17)
     .cmdMoreInfo.Caption = a_bahasa(18)
     .cmdAddVirus.Caption = a_bahasa(19)
     .cmdCancel.Caption = a_bahasa(20)
     
     .cmdClearJail.Caption = a_bahasa(21)
     .cmdKillPris.Caption = a_bahasa(22)
     .cmdRelease.Caption = a_bahasa(23)
     .cmdReleaseTo.Caption = a_bahasa(24)
     
     .cmdCheckUpdate.Caption = j_bahasa(27)
     
End With

With frmRTP
     .cmdFixAllRtp.Caption = a_bahasa(25)
     .cmdFixRtp.Caption = a_bahasa(26)
     .cmdIgnore.Caption = a_bahasa(27)
     .cmdTutup.Caption = a_bahasa(28)
End With

With frmMain
     .TabMain.GantiJudul 1, b_bahasa(0)
     .TabMain.GantiJudul 2, b_bahasa(1)
     .TabMain.GantiJudul 3, b_bahasa(2)
     .TabMain.GantiJudul 4, b_bahasa(3)
     .TabMain.GantiJudul 5, b_bahasa(4)
     
     .TabConfig.GantiJudul 1, b_bahasa(5)
     .TabConfig.GantiJudul 2, b_bahasa(6)
     .TabConfig.GantiJudul 3, b_bahasa(7)
     .TabConfig.GantiJudul 4, b_bahasa(8)
     .TabConfig.GantiJudul 5, b_bahasa(9)
     
     .TabTool.GantiJudul 1, b_bahasa(11)
     .TabTool.GantiJudul 2, b_bahasa(12)
     .TabTool.GantiJudul 3, b_bahasa(13)
     
     .TabAbout.GantiJudul 1, b_bahasa(14)
     .TabAbout.GantiJudul 2, b_bahasa(15)
     
     .FrConfigScan.Caption = c_bahasa(0)
     .FrConfigApp.Caption = c_bahasa(1)
     .frProses.Caption = c_bahasa(2)
     .frModule.Caption = c_bahasa(3)
     .frVirus.Caption = c_bahasa(4)
     .frTemp.Caption = c_bahasa(5)
     .frJail.Caption = c_bahasa(6)
     .frInteralMalware = c_bahasa(7)
     .frSoftInformation = c_bahasa(8)
     .frInternalVirus = c_bahasa(9)
     
     .lbStatus1.Caption = d_bahasa(0)
     .lblProcessed.Caption = d_bahasa(1)
     .lbTime1.Caption = d_bahasa(2)
     .lbFileFound1.Caption = d_bahasa(3)
     .lbFileCheck1.Caption = d_bahasa(4)
     .lbBypass1.Caption = d_bahasa(5)
     .lbMalware1.Caption = d_bahasa(6)
     .lbHidden1.Caption = b_bahasa(3) ' ambil orang lain
     .lbRegistry1.Caption = d_bahasa(7)
     .lblInfor.Caption = d_bahasa(8)
     .lbStatus.Caption = d_bahasa(9)
     .lbMalware.Caption = ": 000000 " & d_bahasa(38)
     .lbHidden.Caption = ": 000000 " & d_bahasa(38)
     .lbReg.Caption = ": 000000 " & d_bahasa(38)
     .lbInfo.Caption = ": 000000 " & d_bahasa(38)

     
     .lblLangSel.Caption = d_bahasa(18)
     .lblLangID.Caption = d_bahasa(19)
     .lblLangAut.Caption = d_bahasa(20)
     
     .lblExceptFolder.Caption = d_bahasa(21)
     .lblExceptFile.Caption = d_bahasa(22)
     .lblExceptReg.Caption = d_bahasa(23)
     .lblAvalaiblePlug.Caption = d_bahasa(24)
     .lblAvalaibleLang.Caption = d_bahasa(25)
     .lblPlugSelect.Caption = d_bahasa(26)
     .lblPlugAut.Caption = d_bahasa(27)
     .lblPlugAutEmail.Caption = j_bahasa(52)
     .lblPlugAutSite.Caption = j_bahasa(53)
     .lblPlugVer.Caption = j_bahasa(54)
     .lblPlugDesc.Caption = d_bahasa(28)
     .lblMalwarePath.Caption = d_bahasa(29)
     .lblMalwareName.Caption = d_bahasa(30)
     
     .lbInfo1(0).Caption = d_bahasa(31)
     .lbInfo1(1).Caption = d_bahasa(32)
     .lbInfo1(2).Caption = d_bahasa(33)
     .lbInfo1(3).Caption = d_bahasa(34)
     .lbInfo1(4).Caption = d_bahasa(35)
     .lbInfo1(5).Caption = d_bahasa(36)
     .lbInfo1(6).Caption = d_bahasa(37)
     .lblLangUsed.Caption = j_bahasa(36)
     
     .mnCScan.Caption = g_bahasa(0)
     .mnEPro.Caption = g_bahasa(2)
     .mnRun.Caption = g_bahasa(3)
     .mnExit.Caption = g_bahasa(4)
     .mnUpdate.Caption = j_bahasa(27)
     
     .mnFixS.Caption = g_bahasa(5)
     .mnFixC.Caption = g_bahasa(6)
     .mnFixA.Caption = g_bahasa(7)
     .mnExcL.Caption = g_bahasa(8)
     .mnExp.Caption = g_bahasa(9)
     .mnProP.Caption = g_bahasa(10)
     
     .mnRefresh.Caption = g_bahasa(11)
     .mnKillPro.Caption = g_bahasa(12)
     .mnRestartPro.Caption = g_bahasa(13)
     .mnPausePro.Caption = g_bahasa(14)
     .mnResumePro.Caption = g_bahasa(15)
     .mnProProperties.Caption = g_bahasa(10)
     
     
     .ck1.Caption = h_bahasa(0)
     .ck2.Caption = h_bahasa(1)
     .ck3.Caption = h_bahasa(2)
     .ck4.Caption = h_bahasa(3)
     .ck5.Caption = h_bahasa(4)
     .ck6.Caption = h_bahasa(5)
     .ck7.Caption = h_bahasa(6)
     .ck8.Caption = h_bahasa(7)
     .ck9.Caption = h_bahasa(8)
     .ck10.Caption = h_bahasa(9)
     .ck11.Caption = h_bahasa(10)
     .ck12.Caption = j_bahasa(50) & " CMC-"

     
End With
End Sub


Private Function ReadExternalLang(sFile As String) As Boolean
Dim IsiFile     As String
Dim SplitterA() As String
Dim SplitterB() As String
Dim SBlok(9)    As String
Dim iCounter    As Long
Dim CutterA     As Long
Dim CutterB     As Long

On Error GoTo LBL_FALSE

If ValidFile(sFile) = False Then GoTo LBL_FALSE
       
IsiFile = ReadUnicodeFile(sFile)
CutterA = InStr(IsiFile, "**/ BEGIN LANG")

' "[Block-0]" : 28
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-0]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(0) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(0), Chr(13))

iCounter = 0
For iCounter = 0 To 28
    SplitterB = Split(SplitterA(iCounter), "=")
    a_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-1]" : 15
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-1]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(1) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(1), Chr(13))

iCounter = 0
For iCounter = 0 To 15
    SplitterB = Split(SplitterA(iCounter), "=")
    b_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-2]" : 09
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-2]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(2) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(2), Chr(13))

iCounter = 0
For iCounter = 0 To 9
    SplitterB = Split(SplitterA(iCounter), "=")
    c_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-3]" : 38
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-3]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(3) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(3), Chr(13))

iCounter = 0
For iCounter = 0 To 38
    SplitterB = Split(SplitterA(iCounter), "=")
    d_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-4]" : 17
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-4]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(4) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(4), Chr(13))

iCounter = 0
For iCounter = 0 To 17
    SplitterB = Split(SplitterA(iCounter), "=")
    e_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------
   
' "[Block-5]" : 21
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-5]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(5) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(5), Chr(13))

iCounter = 0
For iCounter = 0 To 21
    SplitterB = Split(SplitterA(iCounter), "=")
    f_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-6]" : 15
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-6]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(6) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(6), Chr(13))

iCounter = 0
For iCounter = 0 To 15
    SplitterB = Split(SplitterA(iCounter), "=")
    g_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-7]" : 10
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-7]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(7) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(7), Chr(13))

iCounter = 0
For iCounter = 0 To 10
    SplitterB = Split(SplitterA(iCounter), "=")
    h_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-8]" : 27
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-8]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(8) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(8), Chr(13))

iCounter = 0
For iCounter = 0 To 27
    SplitterB = Split(SplitterA(iCounter), "=")
    i_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

' "[Block-9]" : 62
IsiFile = Mid(IsiFile, CutterA)
CutterA = InStr(IsiFile, "[Block-9]") + 10
If CutterA = 0 Then GoTo LBL_FALSE

SBlok(9) = Mid(IsiFile, CutterA)
SplitterA = Split(SBlok(9), Chr(13))

iCounter = 0
For iCounter = 0 To 62
    SplitterB = Split(SplitterA(iCounter), "=")
    j_bahasa(iCounter) = SplitterB(1)
Next
'-----------------------------------

ReadExternalLang = True

Exit Function

LBL_FALSE:
ReadExternalLang = False
End Function



Public Function EnumLangAvalaible(spath As String, LstOut As ListBox) As Boolean
Dim JumLngFile   As Long
Dim iCounter     As Long
Dim StrLngFile() As String
Dim ArrHeadL     As String

JumLngFile = GetFile(spath, StrLngFile)

LstOut.Clear

LstOut.AddItem "=> ENG | Default@1"
LstOut.AddItem "=> INA | Default@2"

For iCounter = 0 To JumLngFile - 1
    ArrHeadL = ReadHeaderLang(StrLngFile(iCounter))
    If Len(ArrHeadL) > 0 Then LstOut.AddItem "=> " & GetLangName(ArrHeadL) & " | " & GetFileName(StrLngFile(iCounter))
Next
End Function

Private Function ReadHeaderLang(sFileLang As String) As String
Dim IsiFile     As String
Dim SplitterA() As String
Dim SplitterB() As String

Dim CutterA As Long
If ValidFile(sFileLang) = False Then Exit Function


IsiFile = ReadUnicodeFile(sFileLang)
CutterA = InStr(IsiFile, "[CMC LANG]")

If CutterA = 0 Then Exit Function
CutterA = InStr(IsiFile, "**/ INIT HEADER INFO") + 21
IsiFile = Mid(IsiFile, CutterA)
SplitterA = Split(IsiFile, Chr(13))

SplitterB = Split(SplitterA(0), "=") 'ID
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"

SplitterB = Split(SplitterA(1), "=") 'NAME
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"

SplitterB = Split(SplitterA(2), "=") 'AUT
ReadHeaderLang = ReadHeaderLang & SplitterB(1) & "\"


End Function

Public Sub WriteLngInfoToLabel(strSelected As String, LBID As Label, LNAME As Label, LAUT As Label)
Dim sTmp        As String
Dim ArrLHead    As String
Dim spath       As String
Dim SplitterA() As String
Dim sNameTmp    As String

' Karena letak filenya namabahasa | namafile
sTmp = Mid(strSelected, InStr(strSelected, "| ") + 2)

Select Case sTmp
    Case "Default@1": sNameTmp = "Default@1": GoTo LBL_KECUALI
    Case "Default@2": sNameTmp = "Default@2": GoTo LBL_KECUALI
End Select

spath = GetFilePath(App_FullPathW(False)) & "\lang"
spath = spath & "\" & sTmp

If ValidFile(spath) = False Then GoTo LBL_FALSE

ArrLHead = ReadHeaderLang(spath)

SplitterA = Split(ArrLHead, "\")

LBID.Caption = ": " & SplitterA(0)
LNAME.Caption = ": " & SplitterA(1)
LAUT.Caption = ": " & SplitterA(2)

Exit Sub

LBL_FALSE:
    LBID.Caption = ": -"
    LNAME.Caption = ": -"
    LAUT.Caption = ": -"

Exit Sub

LBL_KECUALI:
    LBID.Caption = ": Built-in"
    LNAME.Caption = ": " & sNameTmp
    LAUT.Caption = ": A.M Hirin"
End Sub

Private Function GetLangName(LangHead As String) As String
Dim SplitA()  As String

SplitA = Split(LangHead, "\")

GetLangName = SplitA(1)
End Function



'---- Untuk dipakai saat load config
Public Function getNameLangFromFile(sFileLangToRead As String) As String
On Error GoTo LBL_DEF
    getNameLangFromFile = GetLangName(ReadHeaderLang(sFileLangToRead))
     
Exit Function

LBL_DEF:
 getNameLangFromFile = "Default@n"
End Function
