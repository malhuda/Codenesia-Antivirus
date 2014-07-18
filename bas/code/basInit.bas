Attribute VB_Name = "basInit"
' ########################################################
' Module untuk penanganan inisialisasi awal
'
'
Dim cImgMal  As New gComCtl


Public Sub InitAplikasi()

    ' init ektensi penjara
    JailExt = "." & ChrW$(&H6BC) & ChrW$(&H695) & ChrW$(&H6BE)
    ' ini folder jail
    FolderJail = Left(GetSpecFolder(WINDOWS_DIR), 3) & ChrW$(&H6DE) & "~CMC~" & ChrW$(&H6DE)
    'Init Var BERHENTI
    BERHENTI = True

With frmMain
    Call LoadConfig(GetFilePath(App_FullPathW(False)) & "\CMC.ini") ' load config
    Call InitLanguange(LangUsed)
    Call BuildListView
    Call BacaDatabase
    Call LoadDataIcon
    Call InitPHPattern
    Call ENUM_PROSES(.lvProses, frmMain.picBuffer)
    Call BuatPenjara
    Call READ_DATA_JAIL(FolderJail)
    Call ListVirus(.lstListWorm)
    Call BuilDirTree
    Call EnumPlugin(GetFilePath(App_FullPathW(False)) & "\plugin", .lstPlugin)
    Call EnumLangAvalaible(GetFilePath(App_FullPathW(False)) & "\lang", .lstLanguage)
    
    JumPathExcep = ReadExceptPath(GetFilePath(App_FullPathW(False)) & "\Path.lst", .lstExceptFolder)
    JumFileExcep = ReadExceptFile(GetFilePath(App_FullPathW(False)) & "\File.lst", .lstExceptFile)
    JumRegExcep = ReadExceptReg(GetFilePath(App_FullPathW(False)) & "\Reg.lst", .lstExceptReg)
    
 
End With
End Sub


Public Sub BuilDirTree()
If Left$(Command, 2) = "-S" Then
   frmMain.DirTree.LoadTreeDir False, False
   RegNode = False
   StartUpNode = False
   ProsesNode = False
Else
   frmMain.DirTree.LoadTreeDir True, False
   RegNode = True
   StartUpNode = True
   ProsesNode = True
End If
End Sub

' Init sesudah tampilan muncul
Public Sub InitAplikasi2()
Dim nFS As Long
    
    Call UpdateIcon(frmMain.Icon, "CMC PH#3.5 - CodenesiaSoft", frmMain)
    nFS = EnumFileSystem
    
    If nFS = 0 Then
       TampilkanBalon frmMain, i_bahasa(23) & " !", i_bahasa(27), NIIF_WARNING
       Sleep 2000
    End If
    
    If IsWinXPOS = False Then
       If frmMain.ck3.value = 1 Then TampilkanBalon frmMain, j_bahasa(40) & " : " & h_bahasa(2) & " !", i_bahasa(27), NIIF_WARNING
    End If


End Sub
Private Sub BuildListView()

With frmMain
     With .lvMalware ' Listview Malware
          .Font.FaceName = "Arial"
          .Columns.Add , e_bahasa(0), , , lvwAlignCenter, 2200
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 5000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2200
          
                     
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)

     End With
     With .lvRegistry ' Listview Registry
          .Font.FaceName = "Arial"
          .Columns.Add , e_bahasa(4), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(5), , , lvwAlignLeft, 7000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1000
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 3000
          
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With
     With .lvHidden ' Listview Hidden
          .Font.FaceName = "Arial"
          .Columns.Add , e_bahasa(6), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 4000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2000
          
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With
     With .lvInfo ' Listview Information
          .Font.FaceName = "Tahoma"
          .ColorFore = vbBlue
          .Columns.Add , e_bahasa(7), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 4000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 3000
          
          ' Init image list
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With
     With .lvProses ' Listview Proses
          .Font.FaceName = "Arial"
          .Columns.Add , e_bahasa(8), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(9), , , lvwAlignCenter, 1200
          .Columns.Add , "PID", , , lvwAlignCenter, 1200
          .Columns.Add , e_bahasa(10), , , lvwAlignCenter, 1300
          .Columns.Add , e_bahasa(12), , , lvwAlignLeft, 1200
          .Columns.Add , e_bahasa(13), , , lvwAlignLeft, 1300
          .Columns.Add , e_bahasa(14), , , lvwAlignLeft, 1300
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(11), , , lvwAlignLeft, 5000
     End With
     With .lvJail
          .Font.FaceName = "Arial"
          .Columns.Add , e_bahasa(15), , , lvwAlignLeft, 2000
          .Columns.Add , e_bahasa(16), , , lvwAlignLeft, 4000
          .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1300
          .Columns.Add , e_bahasa(17), , , lvwAlignRight, 1600
          
          Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
     End With
End With

With frmRTP.lvRTP
     .Font.FaceName = "Tahoma"
     .Columns.Add , e_bahasa(0), , , lvwAlignCenter, 1800
     .Columns.Add , e_bahasa(1), , , lvwAlignLeft, 4000
     .Columns.Add , e_bahasa(2), , , lvwAlignRight, 1200
     .Columns.Add , e_bahasa(3), , , lvwAlignLeft, 2000
     
     Set .ImageList = cImgMal.NewImageList(16, 16, imlColor32)
End With
     InitImageList
End Sub

Private Sub InitImageList()
With frmMain
     With .lvMalware
          .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic6.hdc, 16, 16
          .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
          .ImageList.AddFromDc frmMain.picCaution.hdc, 16, 16
    End With
    With .lvRegistry
         .ImageList.AddFromDc frmMain.pic7.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic8.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic9.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic10.hdc, 16, 16
    End With
    With .lvHidden
         .ImageList.AddFromDc frmMain.pic11.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic12.hdc, 16, 16
         .ImageList.AddFromDc frmMain.picFolHid.hdc, 16, 16
         .ImageList.AddFromDc frmMain.picFileHid.hdc, 16, 16
         .ImageList.AddFromDc frmMain.picCaution.hdc, 16, 16
    End With
    With .lvInfo
         .ImageList.AddFromDc frmMain.pic14.hdc, 16, 16
         .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
    End With
    With .lvJail
         .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
    End With
End With

With frmRTP.lvRTP
     .ImageList.AddFromDc frmMain.pic1.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic2.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic3.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic4.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic5.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic6.hdc, 16, 16
     .ImageList.AddFromDc frmMain.pic13.hdc, 16, 16
     .ImageList.AddFromDc frmMain.picCaution.hdc, 16, 16
End With
End Sub


