Attribute VB_Name = "basVirus"
' Segala sesuatu tentang virus.... :D

' Kalo mau pake fungsi-fungsi disini harus yakinkan Valid PE

' Memperkenalkan Heuristic Baru atau Ceksum Untuk Virus
' Namanya HBI LX - Heuristical Byte Identification LX - saya pake 30 sementara
' Namanya ngawur ja :D, X menandakan panjang byte untuk sample ceksum virus

Public TmpCeksumPE As String ' menampung ceksum PE
' -- local ajah sama kaya di basPE
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByVal pv6432_lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal pv_lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByVal pv_lpNumberOfBytesRead As Long, ByVal pv_lpOverlapped As Long) As Long
Private Declare Sub RtlZeroMemory Lib "ntdll.dll" (ByVal pDestBuffer As Long, ByVal nDestLengthToFillWithZeroBytes As Long) '<---reset isi dst yaitu mengisinya dengan bytenumber = 0.


Dim PHVirus(6)           As String
Dim PHNameVirus(6)       As String


Public Sub InitPHPattern() ' Polimorphic Worm Masuk Sini

    ' Worm Poli (depanya ada koma itu wajib)
    PHVirus(0) = ",53,83,EC,44,B8,23,10,40,0,B9,0,0,0,0" ' ini worm tapi Poli ceksumnya
    PHNameVirus(0) = "Win32/Mabezat.A"
  
    ' Mulai Virus
    PHVirus(1) = "60,E8,XX,XX,XX,XX,33,C9,8B,2C,24,90,81,XX,XX,XX,XX,XX,81"
    PHNameVirus(1) = "Win32/Sality.A" ' ini pakai cara kusus, karena membuat section baru, biar lebih baik detektornya

    PHVirus(2) = "60,E8,E6,19,0,0,8B,XX,XX,XX,E8,XX,XX,XX,XX,61,68"
    PHNameVirus(2) = "Chirb@mm"

    
    PHVirus(3) = "52,60,B9,XX,XX,XX,XX,E8,0,0,0,0,5F,4F,66" ' XX,XX,66,XX,XX,XX,XX - BYTE relatifnya ad yang dimasukan, karena ada yang sama terus (mgkin krn 0 kali)
    PHNameVirus(3) = "Win32/Gaelicum.A"
    
    PHVirus(4) = "53,60,83,XX,XX,54,5B,53,E8,XX,XX,XX,XX,33,XX,E8"
    PHNameVirus(4) = "Win32/Downloader.NAE"
    
    PHVirus(5) = "BB,XX,XX,XX,XX,93,E9"
    PHNameVirus(5) = "Win32/Mabezat"
    
    PHVirus(6) = "60,E8,XX,XX,XX,XX,61,E9"
    PHNameVirus(6) = "Win32/Expiro"
End Sub

Public Function CocokanVirusWithPHPattern(ByVal DataEPVirus As String) As String
Dim iCount As Byte
' Worm Poli Dulu
  
  If Left(DataEPVirus, Len(PHVirus(0))) = PHVirus(0) Then ' harus tpt 100%
     CocokanVirusWithPHPattern = "PW:" & PHNameVirus(iCount) ' Prefik PW arinta PoliWorm
     Exit Function
  End If


For iCount = 1 To 4
  If HRInstr(DataEPVirus, PHVirus(iCount), 100) > 0 Then
     CocokanVirusWithPHPattern = PHNameVirus(iCount)
     Exit Function
  End If
Next

For iCount = 5 To 6 ' virus kecil data EP nya
  If HRInstr(Left(DataEPVirus, 30), PHVirus(iCount), 100) > 0 Then
     CocokanVirusWithPHPattern = PHNameVirus(iCount)
     Exit Function
  End If
Next

CocokanVirusWithPHPattern = ""

End Function

' disini sambil menyelam sambil minum air (nanti di fungsi ini say ganti menjadi mendapatkan deretan byte sbanyak 256)
' Public Function GetHIBCeksum(hFilePE As Long, nBased As Long, AddNewHeaderBase0 As Long) As String ' return ke string
' virus juga di cek lgsung disini biar lbih ngebut
Public Function GetDataEP(hFilePE As Long, nBased As Long, AddNewHeaderBase0 As Long) As String ' return ke string
Dim INTH32              As IMAGE_NT_HEADERS_32
Dim ISECH()             As IMAGE_SECTION_HEADER
Dim RetFunct            As Long
Dim nNumberBytesOpsRet  As Long
Dim nSection            As Long
Dim pPhysicEP           As Long
Dim iCount              As Integer
Dim OutData()           As Byte
Dim Sec2(1)             As String ' cadangan deteksi sality dan tanatos
Dim nFisik              As Long
Dim nVirtual            As Long
Dim BiggestSectionOff   As Long
Dim SectionToSize       As Long
Dim OPTurnA             As Long
Dim OPTurnB             As Long
Dim btCPattern()        As Byte
Dim KePEHeur            As Boolean ' apakah layak untuk di proses ke PE Heur/Tidak


Dim StrSecAlman         As String

On Error GoTo LBL_AKHIRI

Call SetFilePointer(hFilePE, AddNewHeaderBase0, 0, 0)  '---Base0. lgsung menuju target
RetFunct = ReadFile(hFilePE, VarPtr(INTH32), Len(INTH32), VarPtr(nNumberBytesOpsRet), 0)

  
  nSection = INTH32.FileHeader.NumberOfSections
  
  If nSection <= 0 Then ' masak 0
     GoTo LBL_AKHIRI
  End If
    
  '---cek sectionheader:
  ReDim ISECH(nSection - 1) As IMAGE_SECTION_HEADER
  Call SetFilePointer(hFilePE, AddNewHeaderBase0 + Len(INTH32), 0, 0) '---Base0. INTH32=248 Bytes, set pointernya
  RetFunct = ReadFile(hFilePE, VarPtr(ISECH(0)), Len(ISECH(0)) * nSection, VarPtr(nNumberBytesOpsRet), 0) ' yang akan dibaca ukuran type section (40bytes) x jumlah section
  For iCount = 0 To nSection - 1
      If (INTH32.OptionalHeader.AddressOfEntryPoint >= ISECH(iCount).VirtualAddress) And (INTH32.OptionalHeader.AddressOfEntryPoint < (ISECH(iCount).VirtualAddress + ISECH(iCount).VirtualSize)) Then
          pPhysicEP = ISECH(iCount).PointerToRawData + (INTH32.OptionalHeader.AddressOfEntryPoint - ISECH(iCount).VirtualAddress)
            '---EP-di-file-fisik-ya ketemu,deh!
          If iCount = nSection - 1 And iCount > 1 Then KePEHeur = True Else KePEHeur = False ' layak untuk ke PE Heur karena EP pada section akhir dan section lebih dari 2
          Call ReadUnicodeFile2(hFilePE, pPhysicEP + 1, nBased, OutData)
          StrSecAlman = TataByte(OutData) ' pinjam variablenya yah
       End If
       
       If iCount > 0 Then
          If ISECH(iCount).PointerToRawData > BiggestSectionOff Then
             BiggestSectionOff = ISECH(iCount).PointerToRawData ' biasanya section terakhir
             SectionToSize = ISECH(iCount).SizeOfRawData
          End If
       Else
            BiggestSectionOff = ISECH(iCount).PointerToRawData ' awalnya baygkan terbesar ada yang pertama
            SectionToSize = ISECH(iCount).SizeOfRawData
       End If
       
   Next
   
   'Sekalian disini ngecek ukuran Real dari EXE :D
   nRealSizePE = BiggestSectionOff + SectionToSize
   
   ' Cek Virus dulu
   GetDataEP = CocokanVirusWithPHPattern(StrSecAlman)
   
   If GetDataEP <> "" Then ' dapat virus
      Exit Function ' gak usah proses lagi
   End If
  
    
    
    nFisik = ISECH(nSection - 1).SizeOfRawData ' ukuran section fisik terakhir

   ' Jika OP Code EP pertama adalah &H60 : PUSHAD
   If Left(StrSecAlman, 3) = ",60" Then
      If (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' pastikan sectiony Executable
         'Mainkan sality Awal dulu
         If HRInstr(StrSecAlman, PHVirus(1), 100) > 0 Then
            GetDataEP = PHNameVirus(1)
            Exit Function
         End If
         
         Call ReadUnicodeFile2(hFilePE, ISECH(nSection - 1).PointerToRawData + 1, ISECH(nSection - 1).SizeOfRawData, btCPattern)
         ' Cek Tanatos.M virus poly morphic banyak sampah gak cuma di section, tapi di luar section jg (biar lambat kali cekernya)
              OPTurnA = 0
              For OPTurnA = 0 To (ISECH(nSection - 1).SizeOfRawData - 1)
                    If btCPattern(OPTurnA) = &H8A Then '---8A 44 05 00 = MOV AL,BYTE PTR SS:[EBP+EAX]
                        If btCPattern(OPTurnA + 1) = &H44 Then
                            If btCPattern(OPTurnA + 2) = &H5 Then
                                If btCPattern(OPTurnA + 3) = &H0 Then
                                    OPTurnB = 0 '---preset.
                                    For OPTurnB = (OPTurnA + 4) To (ISECH(nSection - 1).SizeOfRawData - 1)
                                        If btCPattern(OPTurnB) = &H30 Then '---30 07 = XOR BYTE PTR DS:[EDI],AL
                                            If btCPattern(OPTurnB + 1) = &H7 Then
                                                OPTurnB = -1 '---maksimalkan value OPTurnB, yg berarti terlampaui (sudah dapat).
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If OPTurnB = -1 Then
                                        OPTurnA = -1 '---maksimalkan value OPTurnB, yg berarti terlampaui (sudah dapat).
                                        Exit For '---sudah,jangan terlalu lama berputar, 'ntar pusing :)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                
                If OPTurnA = -1 Then ' berhasil dapat Tanatos.M
                   GetDataEP = "Win32/Tanatos.M"
                   Exit Function
                End If
         If nSection > 1 And (ISECH(nSection - 1).Characteristics And &H20) = &H20 Then ' pastikan berisi Code
          ' 2 syarat sudah memenuhi bisa dianggap sality dengan heur
          Sec2(0) = TrimNull0(StrConv(ISECH(1).SectionName(), vbUnicode)) ' nama section ke2
          Sec2(1) = TrimNull0(StrConv(ISECH(nSection - 1).SectionName(), vbUnicode)) ' nama section terakhir
          nVirtual = ISECH(nSection - 1).VirtualSize ' ukuran virtual
          Call CekKemungkinanSality(Sec2(0), Sec2(1), nFisik, nVirtual)  ' lalu cek kemungkinan sality
         End If
      End If
   End If
   
LBL_ALMAN:
   If nFisik >= 36000 Then ' ukuran section fisik terakhir (alman masuk sini)
      Call ReadUnicodeFile2(hFilePE, (ISECH(nSection - 1).PointerToRawData + 1) + (nFisik - 36865), 12000, OutData) ' dari section terakhir offsetnya, 5000 bytes dari kanan
      StrSecAlman = StrConv(OutData, vbUnicode)
      If CekAlman("µí§¶ýÚÿ×Ðþÿÿ·hþÿÿÿÿÿï¡ùÿÿÿÿÿÿÿÿÿÿÿÿ€", StrSecAlman) = True Then
         GetDataEP = "Win32/Alman.A"
         Exit Function
      Else
         If CekAlman("4xÛ 35‰úPC§ãàn†¡út‚t(ZŠð ÐøÈÔ¯éú²/", StrSecAlman) = True Then 'string alman ke 2
            GetDataEP = "Win32/Alman.B"
            Exit Function
         End If
      End If
   Else
      GetDataEP = ""
   End If

If KePEHeur = True Then
   If (ISECH(nSection - 1).Characteristics And &H20) = &H20 And (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' contain code and executable
      nPEHeurGet = "Suspect.PEHeur.2"
   ElseIf (ISECH(nSection - 1).Characteristics And &H20000000) = &H20000000 Then ' executable
      nPEHeurGet = "Suspect.PEHeur.1"
   ElseIf (ISECH(nSection - 1).Characteristics And &H20) = &H20 Then ' contain code
      nPEHeurGet = "Suspect.PEHeur.1"
   Else
      nPEHeurGet = "" ' bebaskan aj lah
   End If
Else
   nPEHeurGet = ""
End If

Erase OutData
Erase btCPattern
Exit Function

LBL_AKHIRI:
    GetDataEP = ""
    nRealSizePE = 0
End Function

' Hanya bekerja setelah fungsi GetDataEP di proses
Public Function GetRealSizePE() As Long
    GetRealSizePE = nRealSizePE
End Function


' dipertajam ah
Function CekKemungkinanSality(nSec1 As String, nSecAkhir As String, SizeSecAkhirFisik As Long, SizeSecAkhirVirt As Long) As String
nSec1 = HilangkanTitik(nSec1)
nSecAkhir = HilangkanTitik(nSecAkhir)
If Mid(nSecAkhir, 2) = nSec1 And SizeSecAkhirVirt = SizeSecAkhirFisik Then
   Select Case SizeSecAkhirFisik
          Case Is > 60000: CekKemungkinanSality = "70% Suspect Tanatos"
          Case Is > 20000: CekKemungkinanSality = "50% Suspect Tanatos"
          Case Is > 10000: CekKemungkinanSality = "40% Suspect Tanatos"
          Case Else: CekKemungkinanSality = ""
   End Select
ElseIf Mid(nSecAkhir, 2) = nSec1 Then
   If SizeSecAkhirFisik > 60000 Then
      CekKemungkinanSality = "50% Suspect Tanatos"
   ElseIf SizeSecAkhirFisik > 10000 Then
      CekKemungkinanSality = "40% Suspect Tanatos"
   End If
ElseIf SizeSecAkhirVirt = SizeSecAkhirFisik Then
   If SizeSecAkhirFisik > 60000 Then
      CekKemungkinanSality = "50% Suspect Tanatos"
   ElseIf SizeSecAkhirFisik > 10000 Then
      CekKemungkinanSality = "40% Suspect Tanatos"
   End If
ElseIf SizeSecAkhirFisik > 60000 Then
   CekKemungkinanSality = "40% Suspect Tanatos"
ElseIf SizeSecAkhirFisik > 20000 Then
   CekKemungkinanSality = "40% Suspect Tanatos"
Else
   CekKemungkinanSality = ""
End If
nSalityGet = CekKemungkinanSality
End Function

' Hanya bekerja setelah fungsi GetDataEP di proses
Private Function CekAlman(ByRef StrInSect As String, ByRef sDataSection As String) As Boolean
If InStr(sDataSection, StrInSect) > 0 Then
    CekAlman = True
Else
    CekAlman = False
End If
End Function


' Untuk EP
Function TataByte(sByte() As Byte) As String
Dim i As Integer
For i = 1 To UBound(sByte) + 1
    TataByte = TataByte & "," & Hex(sByte(i - 1))
Next
End Function


' Buffer
Private Function TrimNull0(sKar As String) As String
TrimNull0 = Left(sKar, InStr(sKar, Chr(0)) - 1)
End Function

' Buffer untuk suspect sality (hilangkan titik section)
Private Function HilangkanTitik(sKarBertitik As String) As String
If Left(sKarBertitik, 1) = "." Then sKarBertitik = Mid(sKarBertitik, 2)
HilangkanTitik = sKarBertitik
End Function


' InSTR Spesial gak pake Telur (dibuat seakurat dan secepat mungkin)
' Fungsi Balik bukan berarti posisi substring pada deretan Hex, jika pola cocok fungsi akan mnghasilkan lebih>0 dan sebaliknya (untuk optimalisasi aja)

Public Function HRInstr(ByVal DeretanHex As String, SubString As String, nProsenSensitif As Byte) As Long
' contoh 29,C0,FE,08,C0,74,XX,75,XX,EB -> XX wajib diberikan pada pattern walpun polanya sudah dpt panjang tanpa XX
Dim MyPos1    As Integer
Dim MyPos2    As Integer
Dim CutString As String
Dim ByHead    As String
Dim TmpPos    As Integer

' Ambil Header dari SubString seblum byte XX
ByHead = GetByteHeader(SubString)

Do
    MyPos1 = InStr(DeretanHex, ByHead)
    If MyPos1 > 0 Then
        If CocokanPolaPendek(SubString, Mid(DeretanHex, MyPos1)) >= nProsenSensitif Then
           TmpPos = MyPos1
           GoTo BROAD_SUCCES
        End If
    Else
        GoTo BROAD_SUCCES
    End If
    DeretanHex = Mid(DeretanHex, MyPos1 + 3) ' + 3 wajib banget
Loop While MyPos1 > 0

HRInstr = 0
Exit Function

BROAD_SUCCES:
    HRInstr = TmpPos

End Function


Private Function GetByteHeader(DeretanByte As String) As String
    GetByteHeader = Left(DeretanByte, InStr(DeretanByte, "XX") - 1)
End Function

' bagi yang pencocokan ByteHeader sebelum XX sudah cocok panggil fungsi sini (meghasilkan prosentasi)
' dengan ini kita bisa milih berapa prosen pola yang cocok (XX tidakdihiraukan) saipa tahu aja pas masukin polanya terlalu pnjang jadi dengan ini bisa diantisipasi
Private Function CocokanPolaPendek(SubStringPola As String, DeretanHexTerpotong As String) As Long
Dim TmpHexTerpotong As String
Dim SubSplitter()   As String
Dim HexSplitter()   As String
Dim HexCocok        As Byte
Dim LengSub         As Byte
Dim MyCount         As Integer
Dim Penambah        As Byte ' penambah karena byte XX gak dihitung

On Error GoTo LBL_FALSE ' eror ya 0

LengSub = Len(SubStringPola)
TmpHexTerpotong = Left(DeretanHexTerpotong, LengSub)

HexSplitter = Split(TmpHexTerpotong, ",") ' byte deteran hex yang sudah disesuakan ukuranya dengan pola virus
SubSplitter = Split(SubStringPola, ",")

For MyCount = 0 To UBound(SubSplitter)
  If HexSplitter(MyCount) = SubSplitter(MyCount) Then HexCocok = HexCocok + 1
  If SubSplitter(MyCount) = "XX" Then Penambah = Penambah + 1 ' artinya XX gak dihiraukan
Next

CocokanPolaPendek = ((HexCocok + Penambah) / (UBound(SubSplitter) + 1)) * 100 ' prosentasinya ketemu deh

Exit Function
LBL_FALSE:
    CocokanPolaPendek = 0
End Function
