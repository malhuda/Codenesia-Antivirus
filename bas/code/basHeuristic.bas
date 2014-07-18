Attribute VB_Name = "basHeuristic"
Public DataAutorun       As String ' buat nampung data autorun
Public TargetShorcutOnFD As String ' buat nampung data target shorcut

Private Function FindShorcutAndTarget(PathYangDiscan As String, ByRef ShorcutTarget As String) As Boolean
Dim lstFile() As String
Dim nFileX    As Long
Dim nTurn2    As Long
Dim TheFile   As String
    
    nFileX = GetFile(Left(PathYangDiscan, 3), lstFile)
    For nTurn2 = 1 To nFileX
        'If BERHENTI = True Then Exit Sub
        TheFile = lstFile(nTurn2 - 1)
        If ValidFile(TheFile) = True Then
           If UCase(Right(TheFile, 4)) = ".LNK" Then
              ' baca target lgsung out
              TargetShorcutOnFD = UCase(GetTargetLink(TheFile, True))
              If Len(TargetShorcutOnFD) > 3 Then
                 ' cek kalo satu alur adalah virus
                 If Left(UCase(PathYangDiscan), 3) = Left(TargetShorcutOnFD, 3) Then
                    If ValidFile(TargetShorcutOnFD) = True Then ' usahakan ahanya yang aktif saja bioar gak mudah ditipu
                       FindShorcutAndTarget = True
                       ShorcutTarget = TargetShorcutOnFD ' tampung di var ini
                       Exit Function ' selesai
                     End If
                 End If
              End If
           End If
        End If
    DoEvents
    Next
    TargetShorcutOnFD = "XX" ' artinya gak ada LNK file
End Function
' Untuk Check Atribute Hidden
Public Function CheckAttrib(sFile As String, bFolder As Boolean)
Dim NAT      As Long
Dim nIcon    As Long

Dim sType    As String
Dim ObjName  As String
Dim sSize    As String

If bFolder = True Then
    nIcon = 1
    sType = "Folder"
    sSize = "N/A"
Else
    nIcon = 0
    sType = "File"
    sSize = Format$(nSizeGlobal, "#,#")
End If

NAT = GetFileAttributes(StrPtr(sFile))
ObjName = GetFileName(sFile)

If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
    AddInfoToList frmMain.lvHidden, ObjName, sFile, sSize, f_bahasa(0) & " " & sType, nIcon, 18
    nHiddenObj = nHiddenObj + 1
    frmMain.lvHidden.ListItems.Item(nHiddenObj).Cut = True
    frmMain.lbHidden.Caption = ": " & Right$("000000" & nHiddenObj, 6) & " " & d_bahasa(38)
End If

End Function

' --- Heuristic [ArrS 1 dan 2]
Private Function IsArrs(PathFile As String, hFile As Long) As Boolean
Dim nmFile      As String
Dim strDrv      As String
Dim isData      As String
Dim Ukuran      As String
Dim sCeksum     As String
Dim TheSTarget  As String


Dim nDataBase   As Long

On Error GoTo KELUAR

nmFile = Mid(PathFile, 4) ' tanpa drive
strDrv = Left(PathFile, 3) ' drive

'- INGAT ARRS hanya Untuk Removable Drive/FD karena bisa saja virus menipu informasi
If GetDriveType(strDrv) <> 2 Then GoTo KELUAR ' 2 = FD

If ValidFile(strDrv & "Autorun.Inf") = True Then ' ArrS 1
   'DoEvents
   
   If DataAutorun = "" Then ' jika blum baca baca, hemat 1x aja bacanya :)
      isData = ReadUnicodeFile(strDrv & "Autorun.Inf")
      isData = UCase(Replace(isData, Chr(0), "")) ' cuma untuk buffer aj
      If Len(isData) > 300 Then isData = Mid(isData, InStr(isData, "OPEN"))
      DataAutorun = isData
   End If
   
   If InStr(1, DataAutorun, UCase(nmFile), vbTextCompare) > 0 Then
      IsArrs = True
      GoTo LBL_MASUKAN_DATA
   Else
      IsArrs = False
   End If
End If

' Tahap 2 "LNK"
  
If Len(TargetShorcutOnFD) = 0 Then ' cari yang pertama kalinya
   If FindShorcutAndTarget(strDrv, TheSTarget) = True Then
      ' shorcut target satu alur (mgkin virus di FD)
      GoTo LBL_MASUKAN_DATA2
   End If
ElseIf Len(TargetShorcutOnFD) > 2 Then ' mgkin ada targetnya
   If UCase(PathFile) = TargetShorcutOnFD Then
      IsArrs = True
      GoTo LBL_MASUKAN_DATA
   End If
Else ' emang gak ada LNK file (XX)
   IsArrs = False
End If
        
Exit Function ' keluar ampe disini

LBL_MASUKAN_DATA:
    sCeksum = MYCeksum2(PathFile)
    GoTo LBL_PROSES_DATA
    
LBL_MASUKAN_DATA2:
    sCeksum = MYCeksum2(TheSTarget)
    
LBL_PROSES_DATA:
    If sCeksum = vbNullString Or sCeksum = String(Len(sCeksum), "0") Then
       sCeksum = MYCeksumCadangan(PathFile, hFile)
       If sCeksum = vbNullString Then
          Exit Function
       End If
    Exit Function ' artinya ceksumnya gak bisa dibaca alias 0
    End If
    
    JumVirusUser = JumVirusUser + 1
    sMD5User(JumVirusUser - 1) = sCeksum ' masukan ke database sementara [numer database sesuai kepala nilai ceksum] pada indek ke-1 aj
    sNamaVirusUser(JumVirusUser - 1) = "Virus [ArrS Method]"
                                          
    TutupFile hFile

KELUAR:
End Function


' Di nonaktifkan karena udah ada penggantinya -> lihat bas Virus
'----[Heuristic Alman] -> Model Lama
Private Function CheckAlman(Where As String, hFile As Long, nSize As Long) As Boolean
Dim Awal         As Long
Dim Panjang      As Long
Dim OutData()    As Byte
Dim Alman(1)     As String
Dim IsiFile      As String

On Error GoTo KELUAR
Alman(0) = "¯EI5œ‚ÞùWç‘Ï" ' :: Alman A
Alman(1) = "µí§¶ýÚÿ×Ðþÿÿ·hþÿÿÿÿÿï¡ùÿÿÿÿÿÿÿÿÿÿÿÿ" ':: Alman B


If nSize > 40970 Then 'And isValidPE32(hFile) > 0 Then '+ Yakinkan PE
    Awal = nSize - 40000 ' yah sekitar 40KB an aj ambil datanya
    Panjang = 40000
    Call ReadUnicodeFile2(hFile, Awal, Panjang, OutData)
    IsiFile = StrConv(OutData, vbUnicode)

    If InStr(IsiFile, Alman(0)) > 1 Then 'Or InStr(isiFile, Alman(1)) > 1 Then
       CheckAlman = True
       TutupFile hFile ' nutupnya klo TRUE saja
    Else
       CheckAlman = False
    End If
Else
    CheckAlman = False
End If

KELUAR:
End Function


' -- [Heuristic Icon]
Private Function CheckIcon(Where As String, hFile As Long) As Boolean
On Error GoTo KELUAR
If IsPE32EXE = False Then GoTo KELUAR

If DRAW_ICO(Where, frmMain.picTmpIcon) = True Then
    CheckIcon = True
    TutupFile hFile
    Exit Function
Else
    CheckIcon = False
End If
KELUAR:
End Function


Private Function CekVBS(sFile As String, hFile As Long) As Boolean
'On Error Resume Next
Dim JumNumer    As Long
Dim iCount      As Long
Dim JumKar      As Long
Dim MySize      As Long
Dim AscKar      As Byte
Dim Pos_Akhir   As Long

Dim OutData()   As Byte
Dim OutData2()  As Byte

If UCase(Right(sFile, 4)) = ".VBS" Then ' Hanya ektensi VBS
'DoEvents
   MySize = GetSizeFile(hFile)
   If MySize > 9000 Then
      Call ReadUnicodeFile2(hFile, 1, 4500, OutData) ' 4500 dari depan
      Call ReadUnicodeFile2(hFile, MySize - 4500, 4500, OutData2)
      isiVBS = StrConv(OutData, vbUnicode)
      isiVBS = isiVBS & StrConv(OutData2, vbUnicode) ' 4500 dari belakang
      Erase OutData()
      Erase OutData2()
   Else
      Call ReadUnicodeFile2(hFile, 1, MySize, OutData)
      isiVBS = StrConv(OutData, vbUnicode)
      Erase OutData()
   End If
   
   isiVBS = UCase(Replace(isiVBS, Chr(0), "")) ' [pembufferan hilangkan char 0]
   
      If InStr((isiVBS), "AUTORUN") > 0 And InStr((isiVBS), "WSCRIPT") > 0 Then GoSub BENAR
   '---- ENKRIPSI
   Pos_Akhir = Len(isiVBS)
   
   For iCount = 1 To Pos_Akhir
       AscKar = Asc(Mid(isiVBS, iCount, 1))
       If AscKar >= 32 And AscKar <= 57 Then
          JumNumer = JumNumer + 1
       Else
          JumKar = JumKar + 1
       End If
   DoEvents
   Next
   
   If JumNumer > JumKar Then GoSub BENAR

Else
   CekVBS = False
End If

Exit Function

BENAR:
    CekVBS = True
    TutupFile hFile
End Function



' --- Fungsi Akumulasi Cek Heuristic
Public Function CekWithHeuristic(sFile As String, hFile As Long) As Boolean
Dim sUkuran As String
sUkuran = Format$(GetSizeFile(hFile), "#,#")
With frmMain
    If IsArrs(sFile, hFile) = True Then
        ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          
          AddInfoToList .lvMalware, f_bahasa(1) & " ArrS", sFile, sUkuran, f_bahasa(2), 1, 18
        GoTo LBL_INFO
    ElseIf CheckIcon(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
           'sUkuran = Format$(GetSizeFile(hFile), "#,#")
           AddInfoToList .lvMalware, f_bahasa(1) & " Icon Detection", sFile, sUkuran, f_bahasa(2), 4, 18
       GoTo LBL_INFO
    ElseIf CekVBS(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          'sUkuran = Format$(GetSizeFile(hFile), "#,#")
          AddInfoToList .lvMalware, f_bahasa(1) & " VBS Heur", sFile, sUkuran, f_bahasa(2), 3, 18
       GoTo LBL_INFO
    End If
End With

Exit Function

LBL_INFO:
    CekWithHeuristic = True
    VirStatus = True
    VirusFound = VirusFound + 1
    frmMain.lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
End Function

' --- Fungsi Akumulasi Cek Heuristic di RTP
Public Function CekWithHeuristicRTP(ByRef sFile As String, ByVal hFile As Long) As Boolean
Dim sUkuran As String
sUkuran = Format$(GetSizeFile(hFile), "#,#")
With frmRTP
    If IsArrs(sFile, hFile) = True Then
        ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          AddInfoToList .lvRTP, f_bahasa(1) & " ArrS", sFile, sUkuran, f_bahasa(2), 1, 18
        GoTo LBL_INFO
    ElseIf CheckIcon(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
           AddInfoToList .lvRTP, f_bahasa(1) & " Icon Detection", sFile, sUkuran, f_bahasa(2), 4, 18
       GoTo LBL_INFO
    ElseIf CekVBS(sFile, hFile) = True Then
       ' klo pngecualian keluar
       If ApaPengecualianFile(sFile, JumFileExcep) = True Then Exit Function
          'sUkuran = Format$(GetSizeFile(hFile), "#,#")
          AddInfoToList .lvRTP, f_bahasa(1) & " VBS Heur", sFile, sUkuran, f_bahasa(2), 3, 18
       GoTo LBL_INFO
    End If
End With

Exit Function

LBL_INFO:
    CekWithHeuristicRTP = True
End Function

' Fungsi Mengenai segala sesuatu yang perlu di informasikan ke User (klo bisa informasi hanya untuk executable file exe,scr,com, tapi bgaiaman cara filter melalui charaterisnya)
' DLL- standar banyak yang bermuatan byte tambahan gak jelas gitu
Public Function CekInformation(ByRef sFile As String, hFile As Long) ' jika di RTP suspect sality ditampilkan
Dim nSizeTmp    As Long
Dim nSize       As Long
Dim TheExt      As String


With frmMain
TheExt = UCase(Right(sFile, 3))
If GetDriveType(Left(sFile, 3)) = 2 Then ' jika di FD
   If TheExt = "DLL" Or TheExt = "EXE" Or TheExt = "SYS" Or TheExt = "OCX" Then
   Else
      InfoFound = InfoFound + 1 ' un proper ektensi di FD kita masukan ke Info
      nSize = GetSizeFile(hFile)
      .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
      AddInfoToList .lvInfo, "Unproper PE Extension", sFile, Format$(nSize, "#,#"), "Perhatian : Waspadai sedikit file ini !", 0, 18
      Exit Function
   End If
End If

If GetRealSizePE = 0 Or IsPE32EXE = False Then Exit Function ' harus benar-benar EXE (bukan dll/sys)

nSize = GetSizeFile(hFile)
   
    If nSalityGet <> "" Then ' ada kemungkinan sality dan varianya (Heur ini belum valid) jadi masuk Informasi
       InfoFound = InfoFound + 1
       .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
       AddInfoToList .lvInfo, nSalityGet, sFile, Format$(nSize, "#,#"), f_bahasa(1) & " Win32/Sality.X", 1, 18
       nSalityGet = "" ' reset
    ElseIf nPEHeurGet <> "" Then ' pengecekan PE Heur
       InfoFound = InfoFound + 1
       .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
       AddInfoToList .lvInfo, nPEHeurGet, sFile, Format$(nSize, "#,#"), f_bahasa(1) & " PE.Heuristic", 1, 18
       nPEHeurGet = "" ' reset
    ElseIf nSize < GetRealSizePE Then ' berari korupt
       InfoFound = InfoFound + 1
       .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
       nSizeTmp = GetRealSizePE - nSize
       AddInfoToList .lvInfo, f_bahasa(3), sFile, Format$(nSize, "#,#"), "Lost " & nSizeTmp & " bytes - MayBey PE cannot work", 0, 18
    ElseIf (nSize - GetRealSizePE) >= GetRealSizePE And nSize > 95000 Then ' ukuranya lebih dari 90 KB-an juga
       InfoFound = InfoFound + 1
       .lbInfo.Caption = ": " & Right$("000000" & InfoFound, 6) & " " & d_bahasa(38)
       nSizeTmp = nSize - GetRealSizePE
       AddInfoToList .lvInfo, f_bahasa(4), sFile, Format$(nSize, "#,#"), nSizeTmp & "B." & f_bahasa(5), 0, 18
    End If

End With

End Function

' Kecurigaan terhdap virus2 masuk ke RTP
Public Function CekInformationRTP(ByRef sFile As String, ByVal hFile As Long) ' jika di RTP suspect sality ditampilkan
Dim szSize As String

szSize = Format$(GetSizeFile(hFile), "#,#")
With frmRTP

If nSalityGet <> "" Then
   AddInfoToList .lvRTP, nSalityGet, sFile, szSize, f_bahasa(1) & " Win32/Sality.X", 2, 8
   nSalityGet = "" ' reset
ElseIf nPEHeurGet <> "" Then
   AddInfoToList .lvRTP, nPEHeurGet, sFile, szSize, f_bahasa(1) & " PE.Heuristic", 2, 8
   nPEHeurGet = "" ' reset
End If

End With

End Function


' Heur untuk cek autorun yang hidden ajh
Public Function CekAutorun(ByRef sRootAutorun As String) As Boolean
Dim lngItem As Long
Dim IsiAR   As String
With frmMain

If UCase(GetFileName(sRootAutorun)) = "AUTORUN.INF" Then
   NAT = GetFileAttributes(StrPtr(sRootAutorun))
   If (NAT = 2 Or NAT = 34 Or NAT = 3 Or NAT = 6 Or NAT = 22 Or NAT = 18 Or NAT = 50 Or NAT = 19 Or NAT = 35) Then
      ' yang di true status hidden aj
   ' klo pngecualian keluar
   If ApaPengecualianFile(sRootAutorun, JumFileExcep) = True Then Exit Function
   
      AddInfoToList .lvMalware, "Suspected ! [Autorun]", sRootAutorun, Format$(FileLen(sRootAutorun), "#,#"), f_bahasa(2) & " Malware Runner", 1, 18
      VirusFound = VirusFound + 1
      
      .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
      
      CekAutorun = True
   Else
      IsiAR = ReadUnicodeFile(sRootAutorun)
      IsiAR = UCase(Replace(IsiAR, Chr(0), ""))
      If InStr(IsiAR, "WSCRIPT.EXE") > 0 Then
         AddInfoToList .lvMalware, "Suspected ! [Autorun]", sRootAutorun, Format$(FileLen(sRootAutorun), "#,#"), f_bahasa(2) & " Malware Runner", 1, 18
         VirusFound = VirusFound + 1
      
        .lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)

         CekAutorun = True
      Else
         CekAutorun = False
      End If
   End If
Else
   CekAutorun = False
End If

End With
End Function


' Heur untuk cek *.lnk yang kemungkinan virus
' Target Dibaca
Public Function CeklnkFolder(ByRef sPathFile As String) As Boolean
Dim lnkString(1) As String ' semntara smplenya baru dua
Dim TheTarget2   As String
Dim nTurn        As Long
Dim MyHnd        As Long

lnkString(0) = Chr(13) & ".com" ' suspect Lnk ke virus lain
lnkString(1) = "wscript.exe" ' ke VBS

TheTarget2 = GetFileName(GetTargetLink(sPathFile, False))
       If Len(TheTarget2) > 0 Then ' format LNK true
           For nTurn = 0 To 1
               If LCase(TheTarget2) = lnkString(nTurn) Then ' ada
                  ' klo pngecualian keluar
                  If ApaPengecualianFile(sPathFile, JumFileExcep) = True Then Exit Function
                  CeklnkFolder = True
                  AddInfoToList frmMain.lvMalware, "Suspect ! [Mal-Shortcut]:" & nTurn, sPathFile, "N/A", f_bahasa(2) & " Malware Runner", 1, 18
                  ' skarang cek isi dalam shorcutnya
                  GoTo lbl_cek_isi ' cek isi hanya pada kasus True saja (mengurangi false detek)
               End If
            Next
Exit Function
lbl_cek_isi:
           TheTarget2 = GetTargetLink(sPathFile, True)
           If GetDriveType(Left(TheTarget2, 3)) <> 2 Then Exit Function ' ingat hanya di FD aj
           If Left(TheTarget2, 3) = Left(sPathFile, 3) Then ' artinya satu jalur
              MyHnd = GetHandleFile(TheTarget2)
              If MyHnd > 0 Then
                 AddInfoToList frmMain.lvMalware, f_bahasa(1) & " ArrS", TheTarget2, Format$(GetSizeFile(MyHnd), "#,#"), f_bahasa(2), 1, 18
                 TutupFile MyHnd
                 KunciFile TheTarget2
                 VirusFound = VirusFound + 1 ' dapat doble
                 frmMain.lbMalware.Caption = ": " & Right$("000000" & VirusFound, 6) & " " & d_bahasa(38)
                 Exit Function
              End If
           End If
       Else
           CeklnkFolder = False
       End If
End Function

' Heur untuk cek *.lnk yang kemungkinan virus
' Target Dibaca - Untuk RTP
Public Function CeklnkFolderRTP(ByRef sPathFile As String) As Boolean
Dim lnkString(1) As String ' semntara smplenya baru dua
Dim TheTarget2    As String
Dim nTurn        As Long
Dim MyHnd        As Long

lnkString(0) = Chr(13) & ".com" ' suspect Lnk ke virus lain
lnkString(1) = "wscript.exe" ' ke VBS

TheTarget2 = GetFileName(GetTargetLink(sPathFile, False))
       If Len(TheTarget2) > 0 Then ' format LNK true
           For nTurn = 0 To 1
               If LCase(TheTarget2) = lnkString(nTurn) Then ' ada
                  ' klo pngecualian keluar
                  If ApaPengecualianFile(sPathFile, JumFileExcep) = True Then Exit Function
                  CeklnkFolderRTP = True
                  AddInfoToList frmRTP.lvRTP, "Suspect ! [Mal-Shortcut]:" & nTurn, sPathFile, "N/A", f_bahasa(2) & " Malware Runner", 1, 18
                  GoTo lbl_cek_isi ' cek isi hanya pada kasus True saja (mengurangi false detek)
               End If
           Next
Exit Function
lbl_cek_isi:
           TheTarget2 = GetTargetLink(sPathFile, True)
           If GetDriveType(Left(TheTarget2, 3)) <> 2 Then Exit Function ' ingat hanya di FD aj
           If Left(TheTarget2, 3) = Left(sPathFile, 3) Then ' artinya satu jalur
              MyHnd = GetHandleFile(TheTarget2)
              If MyHnd > 0 Then
                 AddInfoToList frmRTP.lvRTP, f_bahasa(1) & " ArrS", TheTarget2, Format$(GetSizeFile(MyHnd), "#,#"), f_bahasa(2), 1, 18
                 TutupFile MyHnd
                 KunciFile TheTarget2
                 Exit Function
              End If
           End If
       Else
           CeklnkFolderRTP = False
       End If
End Function
