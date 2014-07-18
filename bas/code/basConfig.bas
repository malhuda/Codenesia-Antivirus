Attribute VB_Name = "basCOnfig"

Public Sub SaveConfig(sPathFile As String)
Dim IsiFile As String
HapusFile sPathFile
With frmMain
    IsiFile = "[SCAN]" & Chr(13)
    IsiFile = IsiFile & "001-" & .ck1.value & Chr(13)
    IsiFile = IsiFile & "002-" & .ck2.value & Chr(13)
    IsiFile = IsiFile & "003-" & .ck3.value & Chr(13)
    IsiFile = IsiFile & "004-" & .ck4.value & Chr(13)
    IsiFile = IsiFile & "005-" & .ck5.value & Chr(13)
    IsiFile = IsiFile & "006-" & .ck6.value & Chr(13)

    IsiFile = IsiFile & "[APP]" & Chr(13)
    IsiFile = IsiFile & "001-" & .ck7.value & Chr(13) ' autorun
    IsiFile = IsiFile & "002-" & .ck8.value & Chr(13) 'RTP
    IsiFile = IsiFile & "003-" & .ck9.value & Chr(13) ' Online Update
    IsiFile = IsiFile & "004-" & .ck10.value & Chr(13) ' Auto Scan FD
    IsiFile = IsiFile & "005-" & .ck11.value & Chr(13) ' Form Ontop
    IsiFile = IsiFile & "006-" & .ck12.value & Chr(13) ' Contesk Menu
    
    IsiFile = IsiFile & "[LANG]" & Chr(13)
    IsiFile = IsiFile & "001-" & LangUsed
End With

WriteFileUniSim sPathFile, IsiFile

End Sub

Public Sub LoadConfig(spath As String)
Dim IsiFile      As String
Dim SplitIsi()   As String
Dim SplitPart()  As String
Dim lngValue(13) As Long

On Error Resume Next
IsiFile = ReadUnicodeFile(spath)
SplitIsi = Split(IsiFile, Chr(13))

With frmMain

For iCounter = 0 To 13
    SplitPart = Split(SplitIsi(iCounter), "-")
    lngValue(iCounter) = CLng(SplitPart(1))
Next

.ck1.value = lngValue(1)
.ck2.value = lngValue(2)
.ck3.value = lngValue(3)
.ck4.value = lngValue(4)
.ck5.value = lngValue(5)
.ck6.value = lngValue(6)

' ada penyekat disini
.ck7.value = lngValue(8) ' autorun
If .ck7.value = 1 Then
   Call InstalInReg(App_FullPathW(False), " -A")
   .mnRun.Checked = True
Else
   Call UnInstalInReg("C.M.C+")
   .mnRun.Checked = False
End If

.ck8.value = lngValue(9) ' RTP
If .ck8.value = 1 Then
   StatusRTP = True
   .mnEPro.Checked = True
Else
   StatusRTP = False
   .mnEPro.Checked = False
End If

.ck9.value = lngValue(10) ' Online Update

.ck10.value = lngValue(11) 'USB Detect
If .ck10.value = 1 Then
   Call GetLasFDVolume
   .tmFlash.Enabled = True
Else
   .tmFlash.Enabled = False
End If

.ck11.value = lngValue(12) ' Form Ontop
If .ck11.value = 1 Then Call LetakanForm(frmMain, True) Else Call LetakanForm(frmMain, False)

.ck12.value = lngValue(13) ' Contek Menu
'If .ck12.Value = 1 Then Install_CMenu (j_bahasa(51) & " CMC") Else unInstall_CMenu (j_bahasa(51) & " CMC")

SplitPart = Split(SplitIsi(15), "-") ' nama bahasanya
LangUsed = SplitPart(1)
.lblLangUsed1.Caption = ": " & getNameLangFromFile(GetFilePath(App_FullPathW(False)) & "\lang\" & LangUsed)

End With
End Sub
