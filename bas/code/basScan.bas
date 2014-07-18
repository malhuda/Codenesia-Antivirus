Attribute VB_Name = "basScan"
' ########################################################
' Module untuk penanganan pencarian file
'
'

Option Explicit

Private Const MAX_PATH  As Long = 260
Private Const MAX_BUF   As Long = 512

Private Const FILE_ATTRIBUTE_READONLY = &H1     '
Private Const FILE_ATTRIBUTE_HIDDEN = &H2     '
Private Const FILE_ATTRIBUTE_SYSTEM = &H4     '
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10     'folder.
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20     '
Private Const FILE_ATTRIBUTE_DEVICE = &H40     '
Private Const FILE_ATTRIBUTE_NORMAL = &H80     '
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100     '
Private Const FILE_ATTRIBUTE_SPARSE_FILE = &H200     '
Private Const FILE_ATTRIBUTE_REPARSE_POINT = &H400     '
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800     'terkompres ntfs.
Private Const FILE_ATTRIBUTE_OFFLINE = &H1000     '
Private Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000     'tidak masuk dalam index pencarian file.
Private Const FILE_ATTRIBUTE_ENCRYPTED = &H4000     'enkripsi ntfs.
Private Const FILE_ATTRIBUTE_VIRTUAL = &H10000     'device virtual;

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long 'FILE_ATTRIBUTES
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH '<>MAX_BUF
    cAlternate          As String * 14
End Type

Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Boolean

Private bLagiJalan  As Boolean

Public Sub KumpulkanFile(ByVal szNamaTarget As String, lbFile As Label, bInfo As Boolean, Optional ByVal YangPertama As Boolean = False)
On Error Resume Next
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    
    If YangPertama = True Then
        bLagiJalan = True
    End If
        
    If bLagiJalan = False Then GoTo ERRHD
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
        GoTo ERRHD
    End If
    Do
        If bLagiJalan = False Then Exit Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
        
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then
                FileFound = FileFound + 1
                frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
                If frmMain.ck1.value = 1 Then
                    If isProperFile(szFullPath, "SYS LNK VBE HTM HTT EXE DLL VBS VMX TML .DB COM SCR BAT INF TML CMD TXT PIF MSI") = True Then
                        FileCheck = FileCheck + 1
                        frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
                        VirStatus = False
                        CocokanDataBase szFullPath ' cek virus atau bukan
                      Else
                        FileNotCheck = FileNotCheck + 1
                        frmMain.lbBypass = ": " & Right$("00000000" & FileNotCheck, 8)
                    End If
                Else
                    FileCheck = FileCheck + 1
                    frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
                    CocokanDataBase szFullPath ' cek virus atau bukan
                End If
                If frmMain.ck4.value = 1 And VirStatus = False Then CheckAttrib szFullPath, False
            Else
               If frmMain.ck4.value = 1 Then CheckAttrib szFullPath, True
            End If

            
            If bInfo = True And WithBuffer = True Then
                frmMain.PB1.value = FileFound
            End If
   
        Else
            bIsFolder = False
        End If
        
        If bIsFolder = True Then
            If bLagiJalan = False Then Exit Do
            lbFile.Caption = szFullPath & "\*.*" 'PotongTampilanKar(szFullPath, 75)
            Call KumpulkanFile(szFullPath, lbFile, bInfo, False) 'enumerasi lagi...
        End If
        
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    Loop While NextStack
    
    Call FindClose(hFind)
ERRHD:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Sub

Public Function BufferPath(szNamaTarget As String, Optional ByVal YangPertama As Boolean = False) As Boolean
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String
    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    If YangPertama = True Then
        bLagiJalan = True
    End If
    
    If bLagiJalan = False Then GoTo ERRHD
    
    szNamaTarget = AddSlashW(szNamaTarget)
    hFind = FindFirstFileW(StrPtr(szNamaTarget & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        If bLagiJalan = False Then GoTo KELUAR
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = szNamaTarget & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
       
       'szFullPath
       'eventa disini
        Else
            bIsFolder = False
        End If
        
        If bIsFolder = True Then
           If bLagiJalan = False Then GoTo KELUAR
           FolToScan = FolToScan + 1
           Call BufferPath(szFullPath, False)  'enumerasi lagi...
        Else
           If bLagiJalan = False Then GoTo KELUAR
           If PathIsDirectory(StrPtr(szFullPath)) = 0 Then FileToScan = FileToScan + 1
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    DoEvents
    Loop While NextStack
    
    frmMain.lbStatus.Caption = j_bahasa(37) & " ! [ " & FolToScan & " " & j_bahasa(39) & ", " & FileToScan & " " & j_bahasa(38) & " ]"
    Call FindClose(hFind)
ERRHD:
    If Err.Number > 0 Then
        Err.Clear
    End If
Exit Function
KELUAR:
WithBuffer = False
End Function

' Pengganti GetFile yang pake FSO
Public Function GetFile(spath As String, ArrFile() As String) As Long
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim cNumber         As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

ReDim ArrFile(1000) As String ' max 1001 file

On Error GoTo ERRHD

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    spath = AddSlashW(spath)
    hFind = FindFirstFileW(StrPtr(spath & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = spath & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If ValidFile(szFullPath) = True Then
              ArrFile(cNumber) = szFullPath
              'MsgBox ArrFile(cNumber)
              cNumber = cNumber + 1
           End If
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    
    DoEvents
    Loop While NextStack
        
    GetFile = cNumber
    Call FindClose(hFind)
ERRHD:
    If Err.Number > 0 Then
        Err.Clear
    End If
End Function

Public Sub ScanRTP(ByRef spath As String)
Dim WFD             As WIN32_FIND_DATA
Dim hFind           As Long
Dim NextStack       As Long
Dim zSlash          As String
Dim szFullPath      As String
Dim szFileName      As String
Dim bIsFolder       As Boolean
Dim DOT1        As String
Dim DOT2        As String

On Error GoTo ERRHD

    DOT1 = ChrW$(46)    '"." dos dir$ 1.
    DOT2 = DOT1 & DOT1  '".." dos dir$ 2.
    zSlash = ChrW$(42) '"\"
    spath = AddSlashW(spath)
    hFind = FindFirstFileW(StrPtr(spath & zSlash), VarPtr(WFD))
    If hFind < 1 Then
       GoTo ERRHD
    End If
    Do
        szFileName = TrimNullW(WFD.cFileName) 'hilangkan NullChar paling kanan.
        szFullPath = spath & szFileName
        bIsFolder = ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
    
        If szFileName <> DOT1 And szFileName <> DOT2 Then
           If bIsFolder = False Then CocokanDataBaseRTP szFullPath
        End If
        NextStack = FindNextFileW(hFind, VarPtr(WFD)) 'unicode;
    
    DoEvents
    Loop While NextStack
     
    Call FindClose(hFind)
ERRHD:
    If Err.Number > 0 Then
        Err.Clear
    End If

End Sub
' Scan Semua file yang ada di root drive 2 dan 3
Public Function ScanRootDrive(lblFile As Label)
On Error Resume Next
Dim lstDrive() As String
Dim lstFile()  As String
Dim nDrive  As Long
Dim nFileX  As Long
Dim nTurn   As Long
Dim nTurn2  As Long

nDrive = GetDrive(lstDrive())

For nTurn = 1 To nDrive
    nFileX = GetFile(lstDrive(nTurn), lstFile)
    For nTurn2 = 1 To nFileX
        If BERHENTI = True Then Exit Function
        If ValidFile(lstFile(nTurn2 - 1)) = True Then
           FileFound = FileFound + 1
           FileCheck = FileCheck + 1
           lblFile.Caption = lstFile(nTurn2 - 1)
           If CekAutorun(lstFile(nTurn2 - 1)) = False Then
              ' cek autotun dulu sebelumnya
              CocokanDataBase lstFile(nTurn2 - 1)
           End If
        End If
    DoEvents
    Next
    nTurn2 = 1
Next

Erase lstDrive
Erase lstFile
End Function


Private Function AddSlashW(ByVal StrInW As String) As String 'OK
On Error Resume Next    'tambah "\" di sebelah kanan string unicode.
    If Right$(StrInW, 1) <> ChrW$(92) Then
        AddSlashW = StrInW & ChrW$(92) 'unicode string;
    Else
        AddSlashW = StrInW
    End If
    Err.Clear
End Function

Private Function TrimNullW(ByVal StInpW As String) As String 'OK
On Error Resume Next
Dim AlignW As Long: AlignW = InStr(StInpW, ChrW$(0))
    If AlignW > 0 Then
        TrimNullW = Left$(StInpW, AlignW - 1) 'unicode string;
    Else
        TrimNullW = StInpW
    End If
End Function

Private Function PotongTampilanKar(sKar As String, nLimit As Byte) As String
If Len(sKar) >= nLimit Then PotongTampilanKar = Left$(sKar, nLimit - 30) & "...\" & GetFileName(sKar) Else PotongTampilanKar = sKar
End Function


Public Sub StopKumpulkan()
On Error Resume Next
    bLagiJalan = False
End Sub
