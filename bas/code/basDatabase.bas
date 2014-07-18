Attribute VB_Name = "basDatabase"
' ########################################################
' Module untuk penanganan akses Database
'
'

Public Function BacaDatabase()
Dim sTemp       As String
Dim sTmp()      As String
Dim sTmp2()     As String
Dim SignBack    As String
Dim pisah       As String
Dim ResPath     As String
Dim spath       As String
Dim iCount      As Integer
Dim iTemp       As Integer
Dim iTurn       As Byte

On Error Resume Next ' Redimensi dulu
ReDim JumlahVirus(15) As Long
ReDim JumlahVirusNonPE(15) As Long

JumVirus = 0 'init

ResPath = GetSpecFolder(WINDOWS_DIR)

SignBack = "x.cmc" ' untuk PE 0x.cmc

' Baca DB PE
For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    ' inisialisasi
    iCount = 0
        
    'spath = "E:\VBA\C.M.C PH#3\sign\no enkrip\" & Hex(iTurn) & SignBack
    'EnkripDB spath, 9, "E:\VBA\C.M.C PH#3\sign\" & Hex(iTurn) & SignBack
    'GoTo LEWAT
    
    spath = GetFilePath(App_FullPathW(False)) & "\sign\" & Hex(iTurn) & SignBack
    'spath = "E:\VBA\C.M.C PH#3\sign\" & Hex(iTurn) & SignBack
    sTemp = ReadDatabaseCMC(spath, 9)
    
    pisah = Chr(13)
    
    
    If sTemp = "" Then GoTo LBL_GAWAT ' gagal baca

    sTmp() = Split(sTemp, pisah)
    iTemp = UBound(sTmp())  ' untuk jumlah virus
    For iCount = 1 To iTemp
        sTmp2() = Split(sTmp(iCount), "=")
        sMD5(iTurn, iCount) = Mid(sTmp2(0), 2)
        sNamaVirus(iTurn, iCount) = sTmp2(1)
    Next
    JumlahVirus(iTurn) = iTemp     ' jumlah virus pada dbx
    JumVirus = JumVirus + (iTemp)  ' jumlah virus pada db0-15
LEWAT:
Next

iTurn = 0
SignBack = "z.cmc" ' untuk non PE 0z.cmc


' Baca DB non PE
For iTurn = 0 To UBound(JumlahVirus) ' 0-15
    ' inisialisasi
    iCount = 0
    
    'spath = "E:\VBA\C.M.C PH#3\signx\no enkrip\" & Hex(iTurn) & SignBack
    'EnkripDB spath, 9, "E:\VBA\C.M.C PH#3\signx\" & Hex(iTurn) & SignBack
    'GoTo LEWAT2
    
    spath = GetFilePath(App_FullPathW(False)) & "\signx\" & Hex(iTurn) & SignBack
    'spath = "E:\VBA\C.M.C PH#3\signx\" & Hex(iTurn) & SignBack
    sTemp = ReadDatabaseCMC(spath, 9)
    
    pisah = Chr(13)
    
    
    If sTemp = "" Then GoTo LBL_GAWAT ' gagal baca

    sTmp() = Split(sTemp, pisah)
    iTemp = UBound(sTmp())  ' untuk jumlah virus
    For iCount = 1 To iTemp
        sTmp2() = Split(sTmp(iCount), "=")
        sMD5nonPE(iTurn, iCount) = Mid(sTmp2(0), 2)
        sNamaVirusnonPE(iTurn, iCount) = sTmp2(1)
    Next
    JumlahVirusNonPE(iTurn) = iTemp     ' jumlah virus pada dbx
    JumVirus = JumVirus + (iTemp)  ' jumlah virus pada db0-15
LEWAT2:
Next

frmMain.lbWorm.Caption = ": " & CStr(JumVirus)

Exit Function
LBL_GAWAT: ' klo ada yang gagal baca
    MsgBox j_bahasa(28) & " ( " & Hex(iTurn) & SignBack & " )", vbCritical
    'End
End Function

Public Function SelectDB(ByRef Ceksum As String) As Long
Select Case Left(Ceksum, 1)
    Case "1": SelectDB = 1
    Case "2": SelectDB = 2
    Case "3": SelectDB = 3
    Case "4": SelectDB = 4
    Case "5": SelectDB = 5
    Case "6": SelectDB = 6
    Case "7": SelectDB = 7
    Case "8": SelectDB = 8
    Case "9": SelectDB = 9
    Case "A": SelectDB = 10
    Case "B": SelectDB = 11
    Case "C": SelectDB = 12
    Case "D": SelectDB = 13
    Case "E": SelectDB = 14
    Case "F": SelectDB = 15
    Case "0": SelectDB = 0
End Select
End Function


' Disini pusat pencocokan baik virus, worm, dan informasi
' RTP tidak masuk sini
Public Function CocokanDataBase(ByRef spath As String) As Boolean ' Memakai turboloop based hex [perulanganya DB di hemat]
Dim iCount       As Integer
Dim Ukuran       As String
Dim CeksumFile   As String
Dim CeksumVirus  As String
Dim nDataBase    As Byte
Dim RetPE        As Long
Dim TmpHGlobal   As Long
Dim RetVirus     As String
Dim RetHeur      As Boolean
On Error GoTo LBL_AKHIR

TmpHGlobal = GetHandleFile(spath)

With frmMain

If TmpHGlobal <= 0 Then GoTo LBL_AKHIR

'Cek *.lnk dulu
If UCase(Right(spath, 4)) = ".LNK" Then
    If CeklnkFolder(spath) = True Then
       VirusFound = VirusFound + 1
       VirStatus = True
       .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

RetPE = IsValidPE32(TmpHGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE > 64 Then ' PE - Ternyata DLL kan bisa diinjek juga gitu
    ' Cek dengan Database Virus dulu jika file PE 32 exe
    RetVirus = GetDataEP(TmpHGlobal, 40, RetPE)
    If RetVirus <> "" Then
       ' klo pengecualian keluar
       If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
          VirusFound = VirusFound + 1
          VirStatus = True
          Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
          .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
          If Left(RetVirus, 3) = "PW:" Then ' artinya hanya WormPoli
             AddInfoToList .lvMalware, Mid(RetVirus, 4), spath, Ukuran, f_bahasa(7), 0, 18
          Else
             AddInfoToList .lvMalware, RetVirus, spath, Ukuran, f_bahasa(6), 2, 18
          End If
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

If RetPE > 0 Then ' tergolong PE
       CeksumFile = MYCeksum(spath, TmpHGlobal)
       
       If CeksumFile = String(Len(CeksumFile), "0") Then
          CeksumFile = MYCeksumCadangan(spath, TmpHGlobal)
       End If

       nDataBase = SelectDB(CeksumFile)
    
       'Ceksumer PE
       For iCount = 1 To JumlahVirus(nDataBase)
         If sMD5(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
           VirusFound = VirusFound + 1
           VirStatus = True
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvMalware, sNamaVirus(nDataBase, iCount), spath, Ukuran, f_bahasa(7), 0, 18
           .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
            GoTo LBL_AKHIR
         End If
         DoEvents
      Next
  
Else
   CeksumFile = MYCeksum(spath, TmpHGlobal)
       
   If CeksumFile = String(Len(CeksumFile), "0") Then
      CeksumFile = MYCeksumCadangan(spath, TmpHGlobal)
   End If
   
   nDataBase = SelectDB(CeksumFile)
    
   ' Ceksumer nonPE
    For iCount = 1 To JumlahVirusNonPE(nDataBase)
        If sMD5nonPE(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
           VirusFound = VirusFound + 1
           VirStatus = True
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvMalware, sNamaVirusnonPE(nDataBase, iCount), spath, Ukuran, f_bahasa(7), 0, 18
           .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
            GoTo LBL_AKHIR
        End If
        DoEvents
    Next
End If

If JumVirusUser > 0 Then
   ' Cocokan Dengan User
   RetVirus = CocokanVirusUser(CeksumFile)
   If RetVirus <> "" Then
      ' Klo pngecualian keluar
      If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
      VirusFound = VirusFound + 1
      VirStatus = True
      Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
      AddInfoToList .lvMalware, RetVirus, spath, Ukuran, f_bahasa(7), 0, 18
      .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
      GoTo LBL_AKHIR
    End If
End If

VirStatus = False ' set false


If frmMain.ck2.value = 1 Then RetHeur = CekWithHeuristic(spath, TmpHGlobal)

' jika heuristic tidak nemu, option di set dan PE valid
If RetHeur = False And .ck5.value = 1 And RetPE > 0 Then
   Call CekInformation(spath, TmpHGlobal) ' tidak diakai RTP
Else
   GoTo LBL_AKHIR
End If

End With

TutupFile TmpHGlobal ' jaga-jaga aja

Exit Function
LBL_AKHIR: ' kalo error/udah dapet lngsung akhiri pemindaian file saat ini
    TutupFile TmpHGlobal
    nSalityGet = "" 'biar gak bahaya
End Function


' Menampilkan Daftar Virus yang ada (ingat kondisi dari tab pemicu ini (cmc info) ini jangan sampe terbuka dulu sblum db dibaca)
Public Sub ListVirus(OutObjek As ListBox)
Dim iTemp       As Integer
Dim nDB         As Byte
Dim lngItem     As Integer

On Error Resume Next
With OutObjek
    .Clear

For nDB = 0 To 15 ' jumlah db=16 --> PE
    For iTemp = 1 To JumlahVirus(nDB)
        lngItem = lngItem + 1
        .AddItem sNamaVirus(nDB, iTemp)
    Next
    iTemp = 0 ' reset
Next
nDB = 0
For nDB = 0 To 15 ' jumlah db=16 --> Non PE
    For iTemp = 1 To JumlahVirusNonPE(nDB)
        lngItem = lngItem + 1
        .AddItem sNamaVirusnonPE(nDB, iTemp)
    Next
    iTemp = 0 ' reset
Next

For iTemp = 1 To JumVirus
   .List(iTemp - 1) = Right$("00000" & CStr(iTemp), 5) & " - " & .List(iTemp - 1)
Next

End With
End Sub

Public Function AddVirusTemp(sFile As String, NamaVirus As String) As Boolean
Dim sCeksum     As String
Dim nDB         As Long
Dim nJumVirus   As Long
Dim MyHandle    As Long
Dim FalsCek     As String

MyHandle = GetHandleFile(sFile)

sCeksum = MYCeksum(sFile, MyHandle)
TutupFile MyHandle

FalsCek = String(Len(sCeksum), "0")

' pakai cadangan jika perlu
If FalsCek = sCeksum Or sCeksum = vbNullString Then
   sCeksum = MYCeksumCadangan(sFile, MyHandle)
End If

If ValidFile(sFile) = False Or sCeksum = vbNullString Or sCeksum = FalsCek Then
   AddVirusTemp = False
Else
   JumVirusUser = JumVirusUser + 1
   nJumVirus = JumVirusUser - 1
   ' Lalu tambahkan nama virus dan ceksum ke DB kusus user virus
   sMD5User(nJumVirus) = sCeksum ' masukan ke database sementara
   sNamaVirusUser(nJumVirus) = NamaVirus ' nama virusnya
   AddVirusTemp = True
End If
End Function

' untuk mencocokan User Virus (dipanggil jika JumVirusUser>0)
' ArrS juga masuk sini
Private Function CocokanVirusUser(ByRef MyHash As String) As String
Dim MyCounter As Long

For MyCounter = 1 To JumVirusUser
    If sMD5User(MyCounter - 1) = MyHash Then
       CocokanVirusUser = sNamaVirusUser(MyCounter - 1)
    End If
Next MyCounter
End Function



' Membaca DB CMC di folder (sign) - DecCode harus 9
Private Function ReadDatabaseCMC(sFileDatabBase As String, DecCode As Byte) As String
Dim DataKeluar()   As Byte
Dim SignUkuran     As Long
Dim SizeFDB        As Long
Dim hFileDB        As Long
Dim iCount         As Long
Dim PenampungStr   As String
OpenFileNow sFileDatabBase ' hGlobal handlenya

SizeFDB = nSizeGlobal
hFileDB = hGlobal

If hFileDB > 0 Then
   Call ReadUnicodeFile2(hFileDB, 1, 10, DataKeluar)
   PenampungStr = StrConv(DataKeluar, vbUnicode)
   If Left(PenampungStr, 2) = "PH" Then ' header benar
     SignUkuran = CLng(Mid(PenampungStr, 3, 6))
     If (SizeFDB - 10) = SignUkuran Then ' ukuran data disamakan
        Erase DataKeluar
        Call ReadUnicodeFile2(hFileDB, 11, SignUkuran, DataKeluar)
        For iCount = 0 To UBound(DataKeluar)
            DataKeluar(iCount) = DataKeluar(iCount) Xor DecCode ' dekripsi
        Next
        PenampungStr = StrConv(DataKeluar, vbUnicode)
        ReadDatabaseCMC = PenampungStr
     Else
       ReadDatabaseCMC = "" ' udah gugur
     End If
   Else
     ReadDatabaseCMC = "" ' udah gugur
   End If
   TutupFile hFileDB
Else
   ReadDatabaseCMC = ""
End If
End Function

' Untuk Enkripsi database
Public Sub EnkripDB(sFileEnk As String, EnkCode As Byte, sFileOut As String)
Dim PenampungStr   As String
Dim DataKeluar()   As Byte
Dim SignUkuran     As Long
Dim SizeFDB        As Long
Dim hFileDB        As Long
Dim iCount         As Long

OpenFileNow sFileEnk ' hGlobal handlenya

SizeFDB = nSizeGlobal
hFileDB = hGlobal

If hFileDB > 0 Then
   Call ReadUnicodeFile2(hFileDB, 1, SizeFDB, DataKeluar)
   
   TutupFile hFileDB

   For iCount = 0 To UBound(DataKeluar)
       DataKeluar(iCount) = DataKeluar(iCount) Xor EnkCode ' dekripsi
   Next
   PenampungStr = "PH" & Right("000000" & CStr(SizeFDB), 6) & "%%"
   PenampungStr = PenampungStr & StrConv(DataKeluar, vbUnicode)
   
   Erase DataKeluar
   
   If ValidFile(sFileOut) = True Then HapusFile sFileOut
   
   WriteFileUniSim sFileOut, PenampungStr

End If
End Sub


'...................... COCOKAN TAPI MILIK RTP
Public Function CocokanDataBaseRTP(spath As String) As Boolean ' Memakai turboloop based hex [perulanganya DB di hemat]
Dim iCount       As Integer
Dim Ukuran       As String
Dim CeksumFile   As String
Dim CeksumVirus  As String
Dim nDataBase    As Byte
Dim RetPE        As Long
Dim TmpHGlobal   As Long
Dim RetVirus     As String
Dim RetHeur      As Boolean
On Error GoTo LBL_AKHIR

TmpHGlobal = GetHandleFile(spath)

With frmRTP

If TmpHGlobal <= 0 Then GoTo LBL_AKHIR

'Cek *.lnk dulu
If UCase(Right(spath, 4)) = ".LNK" Then
    If CeklnkFolderRTP(spath) = True Then
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

RetPE = IsValidPE32(TmpHGlobal) ' fungsi balik IsValidPE32 adalah AddresOfNewHeader

If RetPE > 64 Then ' Ternyata DLL jga bisa diinjek virus kan
    ' Cek dengan Database Virus dulu jika file PE 32 exe
    RetVirus = GetDataEP(TmpHGlobal, 40, RetPE)
    If RetVirus <> "" Then
       ' klo pengecualian keluar
       If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
          Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
          If Left(RetVirus, 3) = "PW:" Then ' artinya hanya WormPoli
             AddInfoToList .lvRTP, Mid(RetVirus, 4), spath, Ukuran, f_bahasa(7), 0, 18
          Else
             AddInfoToList .lvRTP, RetVirus, spath, Ukuran, f_bahasa(6), 2, 18
          End If
       GoTo LBL_AKHIR ' akhiri aj
    End If
End If

If RetPE > 0 Then ' tergolong PE
       CeksumFile = MYCeksum(spath, TmpHGlobal)
       
       If CeksumFile = String(Len(CeksumFile), "0") Then
          CeksumFile = MYCeksumCadangan(spath, TmpHGlobal)
       End If
       
       nDataBase = SelectDB(CeksumFile)
    
       'Ceksumer PE
       For iCount = 1 To JumlahVirus(nDataBase)
         If sMD5(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirus(nDataBase, iCount), spath, Ukuran, f_bahasa(7), 0, 18
           GoTo LBL_AKHIR
         End If
         DoEvents
       Next

Else
   CeksumFile = MYCeksum(spath, TmpHGlobal)
   
   If CeksumFile = String(Len(CeksumFile), "0") Then
      CeksumFile = MYCeksumCadangan(spath, TmpHGlobal)
   End If
       
   nDataBase = SelectDB(CeksumFile)
    
   ' Ceksumer nonPE
    For iCount = 1 To JumlahVirusNonPE(nDataBase)
        If sMD5nonPE(nDataBase, iCount) = CeksumFile Then  ' jika virus didapet
           ' klo pngecualian keluar
           If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
           Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
           AddInfoToList .lvRTP, sNamaVirusnonPE(nDataBase, iCount), spath, Ukuran, f_bahasa(7), 0, 18
           GoTo LBL_AKHIR
        End If
        DoEvents
    Next
End If

If JumVirusUser > 0 Then
   ' Cocokan Dengan User
   RetVirus = CocokanVirusUser(CeksumFile)
   If RetVirus <> "" Then
      ' Klo pngecualian keluar
      If ApaPengecualianFile(spath, JumFileExcep) = True Then GoTo LBL_AKHIR
      Ukuran = Format$(GetSizeFile(TmpHGlobal), "#,#") ' ukuran dalam string
      AddInfoToList .lvRTP, RetVirus, spath, Ukuran, f_bahasa(7), 0, 18
      GoTo LBL_AKHIR
    End If
End If

VirStatus = False ' set false


RetHeur = CekWithHeuristicRTP(spath, TmpHGlobal)

' jika heuristic tidak nemu, option di set dan PE valid
If RetHeur = False And RetPE > 0 Then
   Call CekInformationRTP(spath, TmpHGlobal)
Else
   GoTo LBL_AKHIR
End If

End With

TutupFile TmpHGlobal ' jaga-jaga aja

Exit Function
LBL_AKHIR: ' kalo error/udah dapet lngsung akhiri pemindaian file saat ini
    TutupFile TmpHGlobal
    nSalityGet = "" 'biar gak bahaya
End Function

