Attribute VB_Name = "basHeal"
'--> memberikan penangan-penanganan kusus untuk virus kusus

'--- W32/Srigala.A [exe]
'--- MH uploader : Sdr. Revil
'--- Ukuran yang saya terima : 9.7 MB
' Tapi saya yakin ukuran asli worm tsb hanya 183 KB-an, karena sifat worm
' menumpuk (simulasi infeksi) file exe lain sehingga badan worm tidak tetap
' untuk itu pembersihan file exe yang diinfeksi (ditumpuk) saya ini mungkin hanya berlaku
' pada worm srigala sesuai yang dikirm (namun bisa saja global)
' Pembersihana berdasar string akhir "|||||MZ" (7 kar), seharusnya saya pakai pendeteksian Header
' exe terakhir dalam sebuah file (saya belum bisa) mgkin karena pola exe 12 byte awal berbeda-beda,
' sehingga data exe-exe virus (exe lain) yang beradi di atasnya  bisa dipotong dengan baik, udah lah biarin pake cara ini dulu....

Private Const HeaderSrigala = "|||||MZ" ' ||||| = string terakhir dari worm dan file exe yang digabung [Kalo bisa ganti const yang mewakili header exe]
Public Function HealInfeksiSrigala(sFileInfect As String) As Long
' uhuy restore semua exe yang di bundle si srigala
Dim sTmp        As String
Dim sTmp2       As String
Dim sIsiFile    As String
Dim sRevTmp     As String
Dim sPathFile   As String
Dim fKusus(1 To 5)   As String

Dim sNum        As Long
Dim lPos        As Long
Dim lPos2       As Long
Dim lUkuran     As Long
' Eh kalo pada point-point folder kusus jangan direstore tapi di kill langsung
fKusus(1) = GetSpecFolder(USER_STARTUP)
fKusus(2) = GetSpecFolder(WINDOWS_DIR) '--> ah gak usah/iya ya, sial neh virus benr buat bimbang

sPathFile = GetFilePath(sFileInfect) ' klo didalam starup jangan direstore
If UCase(fKusus(1)) = UCase(sPathFile) Or UCase(fKusus(2)) = UCase(Left(sPathFile, 10)) Then
   HapusFile sFileInfect ' hapus aj klo  di dua titik path (+sub path) tsb
   Exit Function
End If

sTmp = ReadUnicodeFile(sFileInfect)
BuatFolder sFileInfect & "_FULL_RESTORE"
Do
    lPos = InStr(sTmp, HeaderSrigala)
    If lPos = 0 Then GoTo LBL_AKHIR
    sTmp = Mid(sTmp, lPos + 5) ' -- potong terus sampai akhir
    lPos2 = InStr(sTmp, HeaderSrigala) ' -- deteksi exe selanjtunya [untuk menentukan ukuran exe sebelumnya]
    If lPos2 > 0 Then ' artinya masih lebih dari 1 exe yang mau ditangani
       lUkuran = lPos2 - 1
    Else
       lUkuran = Len(sTmp)
    End If
    sIsiFile = Mid(sTmp, 1, lUkuran)
    sNum = sNum + 1
    WriteFileUniSim sFileInfect & "_FULL_RESTORE" & "\file_restore_cms_(" & sNum & ").exe", sIsiFile '
    
    DoEvents
Loop While lPos > 0

LBL_AKHIR:
   ' yang file exe terakhir diletakan disebelah aj (mungkin exe-nya si korban)
   ' target hapus dulu

   HapusFile sFileInfect
   CopiFile sFileInfect & "_FULL_RESTORE" & "\file_restore_cms_(" & sNum & ").exe", sFileInfect & "_restore.exe", False
   HealInfeksiSrigala = sNum
End Function
