Attribute VB_Name = "basPluginAkses"
Dim sPluginPath()     As String
Dim sPluginName()     As String
Dim sPluginAuthor()   As String
Dim sPluginAutEmail() As String
Dim sPluginAutSite()  As String
Dim sPluginDesc()     As String



Public Sub EnumPlugin(szFolderPlugin As String, LstOut As ListBox)
Dim nPluginCount            As Long
Dim ArPlugsList()           As PG_PLUGIN_GENERAL_INFORMATION
Dim CDCounter               As Long
Dim HeadS                   As String


    
    HeadS = ": "
    nPluginCount = PgEnumeratePluginFiles(szFolderPlugin, ArPlugsList())
    
    
    LstOut.Clear
    If nPluginCount = 0 Then
        LstOut.AddItem j_bahasa(55) & " www.spectraid.cf"
        Exit Sub
    ElseIf nPluginCount < 0 Then
        Exit Sub
    Else
        ReDim sPluginPath(nPluginCount - 1) As String
        ReDim sPluginName(nPluginCount - 1) As String
        ReDim sPluginAuthor(nPluginCount - 1) As String
        ReDim sPluginAutEmail(nPluginCount - 1) As String
        ReDim sPluginAutSite(nPluginCount - 1) As String
        ReDim sPluginDesc(nPluginCount - 1) As String
        
        With frmMain
             .lblPlugSelect1.Caption = HeadS
             .lblPlugAut1.Caption = HeadS
             .lblPlugAutEmail1.Caption = HeadS
             .lblPlugAutSite1.Caption = HeadS
             .lblPlugVer1.Caption = HeadS
             .lblPlugDesc1.Caption = HeadS
        End With

        For CDCounter = 0 To (nPluginCount - 1)
            sPluginPath(CDCounter) = ArPlugsList(CDCounter).szPluginStartupPathW
            sPluginName(CDCounter) = ArPlugsList(CDCounter).szPluginName
            sPluginAuthor(CDCounter) = ArPlugsList(CDCounter).szPluginAuthor
            sPluginAutEmail(CDCounter) = ArPlugsList(CDCounter).szPluginAuthorEMail
            sPluginAutSite(CDCounter) = ArPlugsList(CDCounter).szPluginAuthorSite
            sPluginDesc(CDCounter) = ArPlugsList(CDCounter).szPluginDescription

            LstOut.AddItem "-> " & ArPlugsList(CDCounter).szPluginStartupPathW
        Next
    End If
    
    Erase ArPlugsList
End Sub

Public Sub RetrievePlugInfo(PlugIndek As Long, LstRead As ListBox, lblPlugName As Label, lblAut As Label, lblAutEmail As Label, lblAutSite As Label, lblVerCode As Label, LblDesc As Label)
Dim HeadS                   As String
Dim hIcon                   As Long

    
    HeadS = ": "
    If PlugIndek >= 0 And ValidFile(sPluginPath(PlugIndek)) = True Then
       lblPlugName.Caption = HeadS & sPluginName(PlugIndek)
       lblAut.Caption = HeadS & sPluginAuthor(PlugIndek)
       lblAutEmail.Caption = HeadS & sPluginAutEmail(PlugIndek)
       lblAutSite.Caption = HeadS & sPluginAutSite(PlugIndek)
       LblDesc.Caption = HeadS & sPluginDesc(PlugIndek)
       
       ' gambar Iconya
       CopiFile sPluginPath(PlugIndek), "C:\$$$$$.exe", False
       RetrieveIcon "C:\$$$$$.exe", frmMain.picPlugin, ricnLarge
       HapusFile "C:\$$$$$.exe"
    End If
    
End Sub

Public Sub RunPlugin(PlugIndek As Long)
Dim RetRun                  As Long
Dim szPluginFileNameW       As String

            If PlugIndek >= 0 Then
               szPluginFileNameW = sPluginPath(PlugIndek)
               If ValidFile(szPluginFileNameW) = False Then Exit Sub
               If MsgBox(j_bahasa(56) & " (" & sPluginName(PlugIndek) & ") ?", vbDefaultButton1 + vbQuestion + vbYesNo) = vbYes Then
                  If MsgBox(j_bahasa(57) & " ?", vbDefaultButton1 + vbYesNo + vbQuestion) = vbYes Then
                     RetRun = PgLoadAndRunPlugin(szPluginFileNameW, True)
                     If RetRun = 0 Then
                        MsgBox j_bahasa(58) & " !", vbExclamation
                     End If
                  Else
                     RetRun = PgLoadAndRunPlugin(szPluginFileNameW, True)
                     If RetRun = 0 Then
                        MsgBox j_bahasa(58) & " !", vbExclamation
                     End If
                  End If
               End If
            End If

End Sub
