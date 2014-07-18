Attribute VB_Name = "basScanSystem"
' Untuk Scan Service, Proses Dan StartUp {registry} + rootdrive

Dim PID_ToTerminated(70)    As Long
Dim PID_ToRestarted(70)     As Long
Dim nTerminate              As Long
Dim nRestart                As Long
Dim nKunci                  As Long
Dim MYID                    As Long
Dim Path_Terminate(70)      As String
Dim Path_Restart(70)        As String
Dim Path_ToKunci(70)        As String

Dim CLFL As New classFile

' BUFFER path output service
Private Function BufferServiceOutPath(sBuf As String) As String
If UCase(Left(sBuf, 12)) = "\SYSTEMROOT\" Or UCase(Left(sBuf, 8)) = "SYSTEM32" Then
   If UCase(Left(sBuf, 12)) = "\SYSTEMROOT\" Then
      BufferServiceOutPath = GetSpecFolder(WINDOWS_DIR) & Mid(sBuf, 12)
   Else
      BufferServiceOutPath = GetSpecFolder(WINDOWS_DIR) & "\" & sBuf
   End If
Else
   BufferServiceOutPath = sBuf
End If
End Function


'-----SCAN SERVICE
Public Sub ScanService(lbFile As Label, bDelete As Boolean)
Dim NSRVO()     As ENUMERATE_SERVICES_OUTPUT
Dim LEAX        As Long
Dim CTRN        As Long
Dim StatDestr   As Long

Dim sNameService    As String
Dim sServicePath    As String

On Error Resume Next


With frmMain.lvMalware
    LEAX = PamzEnumerateServices("", True, True, True, True, NSRVO())
    If LEAX > 0 Then
    ' di init PB1 maximal jumlah service
    frmMain.PB1.Max = LEAX
        For CTRN = 0 To (LEAX - 1)
            sNameService = NSRVO(CTRN).szServiceNameW
            sServicePath = BufferServiceOutPath(NSRVO(CTRN).szServiceApproxPathW)
            lbFile.Caption = j_bahasa(3) & " - " & sServicePath
            If ValidFile(sServicePath) = True Then  ' yakinkan yang discan adalah file
               FileFound = FileFound + 1
               FileToScan = FileToScan + 1
               FileCheck = FileCheck + 1
               CocokanDataBase sServicePath ' cek service apakah virus atau bukan dan ditambah heuristic (jika diset)
               If BERHENTI = True Then Exit For
               If VirStatus = True Then ' jika status virus true
                  If IsFileProtectedBySystem(sServicePath) = False Then
                     StatDestr = PamzDestroyService("", sNameService, bDelete)
                     If StatDestr = 0 Then
                        .ListItems.Item(.ListItems.count).SubItem(4).Text = j_bahasa(1) & " !" ' status diganti
                     Else
                        .ListItems.Item(.ListItems.count).SubItem(4).Text = j_bahasa(2) ' status diganti
                     End If
                  Else ' waduh file system
                     .ListItems.Item(.ListItems.count).IconIndex = 7
                     .ListItems.Item(.ListItems.count).SubItem(4).Text = f_bahasa(20)
                  End If
               End If
            End If
            
            If WithBuffer = True Then
                frmMain.PB1.value = CTRN
            End If
            
        Next
    End If
    Erase NSRVO
End With
End Sub


'----- SCAN STARTUP
Public Sub ScanRegStartup(lbFile As Label, bWithCommon As Boolean)
Dim nJum         As Long
Dim nLong        As Long
Dim nStart       As Long
Dim lngItem      As Long
Dim sFile        As String
Dim sPathReg(7)  As String
Dim sKeyRegN(7)  As String
Dim sValueName() As String
Dim sValueData() As String
Dim sName        As String
Dim NamaVrz      As String

VirStatus = False 'init

sKeyRegN(0) = "HKCU"
sPathReg(0) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(1) = "HKLM"
sPathReg(1) = "Software\Microsoft\Windows\CurrentVersion\Run"

sKeyRegN(2) = "HKLM"
sPathReg(2) = "Software\Microsoft\Windows\CurrentVersion\Run-"

sKeyRegN(3) = "HKLM"
sPathReg(3) = "Software\Microsoft\Windows\CurrentVersion\RunOnce"

sKeyRegN(4) = "HKLM"
sPathReg(4) = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"

With frmMain.lvMalware

For nStart = 0 To 4
    nJum = RegEnumStr(SingkatanKey(sKeyRegN(nStart)), sPathReg(nStart), sValueName(), sValueData())
    For nLong = 1 To nJum
        sFile = BufferStartupPath(sValueData(nLong))
        If ValidFile(sFile) = True Then
           FileFound = FileFound + 1
           FileToScan = FileToScan + 1
           FileCheck = FileCheck + 1
           lbFile.Caption = "[" & sValueName(nLong) & "] - " & sFile
           
           frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
           frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
           
           CocokanDataBase sFile
           If VirStatus = True Then ' Startup adalah virus
              nErrorReg = nErrorReg + 1 ' bilangan reg yang bermaslah dinaikan
              lngItem = .ListItems.count
              .ListItems.Item(lngItem).SubItem(4).Text = j_bahasa(4)
              NamaVrz = .ListItems.Item(lngItem).Text
              ' masukan ke info registry
              AddInfoToList frmMain.lvRegistry, sValueName(nLong), sKeyRegN(nStart) & "\" & sPathReg(nStart) & "\" & sValueName(nLong), Len(sValueData(nLong)), "Startup Virus (" & NamaVrz & ")", 2, 18
           Else ' lakukan pengecekan startup tidak lazim (non PE)
              If CekStartupTakLazim(sFile) = True Then
                 ' masukan ke info registry
                 nErrorReg = nErrorReg + 1 ' bilangan reg yang bermaslah dinaikan
                 AddInfoToList frmMain.lvRegistry, sValueName(nLong), sKeyRegN(nStart) & "\" & sPathReg(nStart) & "\" & sValueName(nLong), Len(sValueData(nLong)), "Kunci startup tak lazim", 2, 18
              End If
           End If
        End If
    Next
    ' hayoo habis dipakai diset ulang
    nLong = 1
    Erase sValueName
    Erase sValueData
Next

' Ditambah Reg Start-Up Singgle
sPathReg(5) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(5) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(5)), sPathReg(5), "Shell")
If UCase(sFile) <> "EXPLORER.EXE" Then ' berarti ada tuh
   sFile = Mid(sFile, InStr(sFile, Chr(32)) + 1)
   If ValidFile(sFile) = True Then
      FileFound = FileFound + 1
      FileToScan = FileToScan + 1
      FileCheck = FileCheck + 1
      lbFile.Caption = "[EXPLORER.EXE] - " & sFile
      
      frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
      frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)

      CocokanDataBase sFile
      If VirStatus = True Then ' Startup virus
         nErrorReg = nErrorReg + 1
         lngItem = .ListItems.count
         .ListItems.Item(lngItem).SubItem(4).Text = j_bahasa(6)
         ' jangan masukan ke info registry (bisa difix reg)
      End If
   End If
End If

sPathReg(6) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Winlogon"
sKeyRegN(6) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(6)), sPathReg(6), "Userinit")
sName = GetSpecFolder(SYSTEM_DIR) & "\userinit.exe"
If UCase(sFile) <> UCase(sName) Then  ' berarti ada tuh
   sFile = Replace(UCase(sFile), UCase(sName) & ",", "")
   sFile = BuangSpaceAwal(sFile)
   If ValidFile(sFile) = True Then
      FileFound = FileFound + 1
      FileToScan = FileToScan + 1
      FileCheck = FileCheck + 1
      lbFile.Caption = "[USERINIT] - " & sFile
      
      frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
      frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)

      CocokanDataBase sFile
      If VirStatus = True Then ' Startup virus
         lngItem = .ListItems.count
         .ListItems.Item(lngItem).SubItem(4).Text = "Found in Winlogon-Startup"
         ' jangan masukan ke info registry (bisa difix reg)
      End If
   End If
End If

sPathReg(7) = "SOFTWARE\microsoft\Windows NT\CurrentVersion\Windows"
sKeyRegN(7) = "HKLM"
sFile = GetStringValue(SingkatanKey(sKeyRegN(7)), sPathReg(7), "Load")
If ValidFile(sFile) = True Then
   FileFound = FileFound + 1
   FileToScan = FileToScan + 1
   FileCheck = FileCheck + 1
   lbFile.Caption = "[LOAD] - " & sFile
   
   frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
   frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)

   CocokanDataBase sFile
   If VirStatus = True Then ' Startup virus
      lngItem = .ListItems.count
      .ListItems.Item(lngItem).SubItem(4).Text = "Found in Winload-Startup"
      ' jangan masukan ke info registry (bisa difix reg)
   End If
End If

If bWithCommon = True Then
   KumpulkanFile GetSpecFolder(USER_STARTUP), lbFile, False, True
   KumpulkanFile GetSpecFolder(ALL_USER_STARTUP), lbFile, False, True
End If

End With

frmMain.lbReg.Caption = ": " & Right$("000000" & nErrorReg, 6) & " " & d_bahasa(38)

End Sub


' ---- SCAN PROSES
Public Sub ScanProses(bModuleScan As Boolean, lbProses As Label)
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim PID             As Long
Dim nSize           As Long
Dim lngItem         As Long

Dim ENPC()          As ENUMERATE_PROCESSES_OUTPUT
Dim sProsesPath     As String
Dim WScript         As String ' jika proses WS script dibunuh dulu

    LEAX = PamzEnumerateProcesses(ENPC())
    MYID = GetCurrentProcessId()
    VirStatus = False 'init
    If LEAX <= 0 Then
        GoTo LBL_TERAKHIR
    End If
    nTerminate = 0 ' init
    WScript = GetSpecFolder(WINDOWS_DIR) & "\System32\wscript.exe"
        
    ' di init PB1 maximal jumlah proses
    frmMain.PB1.Max = LEAX
    
    With frmMain.lvMalware
    For CTurn = 0 To (LEAX - 1)
        sProsesPath = PamzNtPathToUserFriendlyPathW(ENPC(CTurn).szNtExecutablePathW)
        PID = ENPC(CTurn).nProcessID
        lbProses.Caption = "(" & PID & ") - " & sProsesPath
        nSize = ENPC(CTurn).nSizeOfExecutableOpInMemory
        If UCase(WScript) = UCase(sProsesPath) Then ' buat jaga-jaga kalo ada WScript
            PamzTerminateProcess (PID)
            GoTo LANJUT_FOR
        End If
        If ValidFile(sProsesPath) = True Then  ' yakinkan yang discan adalah file
            FileFound = FileFound + 1
            FileToScan = FileToScan + 1
            FileCheck = FileCheck + 1
            
            frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
            frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)

            CocokanDataBase sProsesPath ' cek proses apakah virus atau bukan dan ditambah heuristic (jika diset)
        End If
        If BERHENTI = True Then Exit For
        If VirStatus = True And PID <> MYID Then ' jika status virus true namun dengan catatan bukan proses sendiri
           lngItem = .ListItems.count
           PID_ToTerminated(nTerminate) = PID
           Path_Terminate(nTerminate) = sProsesPath
           PamzSuspendResumeProcessThreads PID, False ' di pause dulu
           nTerminate = nTerminate + 1
           ' ganti status
           .ListItems.Item(lngItem).SubItem(4).Text = j_bahasa(7) ' status diganti
        Else ' modulenya di scan klo bukan proses virus [tapi bModuleScan harus true]
            If bModuleScan = True Then ScanModules PID, lbProses, sProsesPath
        End If
        
        If WithBuffer = True Then
           frmMain.PB1.value = CTurn + 1
        End If
        
LANJUT_FOR:
    Next
    End With
    
    Erase ENPC()
    
    CTurn = 0 'reset
        
    ' saat-nya beraksi secara serempak
    
    For CTurn = 1 To nRestart  ' restart proses2 yang terinfeksi module virus
        KillProses PID_ToRestarted(CTurn - 1), Path_Restart(CTurn - 1), True, False
    Next
    
    CTurn = 0 'reset

    For CTurn = 1 To nTerminate ' terminate lalu kunci proses virus
        KillProses PID_ToTerminated(CTurn - 1), Path_Terminate(CTurn - 1), False, True
    Next
    
    CTurn = 0 'reset
    For CTurn = 1 To nKunci ' kusus untuk module-module yang belum ke ke-kunci [kusus proses udah dikunci di atas]
        KunciFile Path_ToKunci(CTurn - 1) ' gak berhasil pake cdangan dulu smntara
    Next
    
LBL_TERAKHIR:
End Sub

Private Function ScanModules(ByVal TargetPID As Long, lbModule As Label, sProses As String) As Boolean ' akan TRUE jika salah satu module adalah virus, lalu proses akan dimatikan atau restart [karena semntara belum punya fungsi untuk unload dll]
On Error Resume Next
Dim LEAX            As Long
Dim CTurn           As Long
Dim pAddress        As String
Dim sModulePath     As String
Dim ENMC()          As ENUMERATE_MODULES_OUTPUT
    LEAX = PamzEnumerateModules(TargetPID, ENMC)
    If LEAX <= 0 Then ' gagal mendapatkan module
        GoTo LBL_TERAKHIR
    End If
    
    With frmMain.lvMalware
    For CTurn = 0 To (LEAX - 1)
        pAddress = Hex$(CLng(PamzNtPathToUserFriendlyPathW(CStr(ENMC(CTurn).pBaseAddress))))
        sModulePath = PamzNtPathToUserFriendlyPathW(ENMC(CTurn).szNtModulePathW)
        lbModule.Caption = "(0x" & pAddress & ") - " & sModulePath
        If ValidFile(sModulePath) = True Then ' yakinkan yang discan adalah file
            FileFound = FileFound + 1
            FileToScan = FileToScan + 1
            FileCheck = FileCheck + 1
            
            frmMain.lbFileFound = ": " & Right$("00000000" & FileFound, 8)
            frmMain.lbFileCheck = ": " & Right$("00000000" & FileCheck, 8)
            CocokanDataBase sModulePath ' cek module apakah virus atau bukan dan ditambah heuristic (jika diset)
        End If
        If BERHENTI = True Then Exit For
        If VirStatus = True And TargetPID <> MYID Then ' jika status virus true (module) & bukan proses sendiri --> di ganti kalo udah ada fungsi unload dll
           lngItem = .ListItems.count
           Path_Restart(nRestart) = sProses ' masukan alamat file proses-nya
           PID_ToRestarted(nRestart) = TargetPID
           Path_ToKunci(nKunci) = sModulePath ' tambahkan path module untuk dikunci
           PamzSuspendResumeProcessThreads TargetPID, False ' di pause dulu prosesnya
           nRestart = nRestart + 1 ' jumlah yang akan distart dinaikan
           nKunci = nKunci + 1 ' Jumlah yang mau dikunci (kunci module-nya biar gak bisa dijalankan lagi)
           .ListItems.Item(lngItem).SubItem(4).Text = j_bahasa(8) ' status diganti
        End If
    Next
    End With
    Erase ENMC()

LBL_TERAKHIR:
End Function

' ---------------------- FUNGSI BUFFER
' Cek Ketidak laziman Startup (Startup non PE)
Private Function CekStartupTakLazim(FilePathStart As String) As Boolean
Dim hTmp As Long
hTmp = CLFL.VbOpenFile(FilePathStart, FOR_BINARY_ACCESS_READ, LOCK_NONE)
If hTmp > 0 Then
   If GetPE3264Type(hTmp) = 0 Then
      CekStartupTakLazim = True ' startup ko bukan PE, hmm jangan-jangan....
   Else
      CekStartupTakLazim = False
   End If
Else
   CekStartupTakLazim = False
End If
CLFL.VbCloseFile hTmp
End Function

