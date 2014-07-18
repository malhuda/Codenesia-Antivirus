Attribute VB_Name = "basContol"
' ########################################################
' Module untuk penanganan Control thdp aplikasi
'
'

' API Untuk menunda Eksekusi
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' API Untuk mengatur peletakan Form
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'File Prop
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long


' Konstanta peletakan form
Public Const SWP_NOMOVE                  As Long = &H2
Public Const SWP_NOSIZE                  As Long = &H1
Public Const HWND_NOTOPMOST              As Long = -2
Public Const HWND_TOPMOST                As Long = -1
Public Const Flags                       As Long = SWP_NOMOVE Or SWP_NOSIZE


Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Public Sub ShowProperties(Filename As String, OwnerhWnd As Long)
On Error Resume Next
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = App.hInstance
        .lpIDList = 0
    End With
    ShellExecuteEx SEI
End Sub


' Fungsi untuk meletakan Form
Public Function LetakanForm(Frm As Form, Top As Boolean)
If Top = True Then
    Call SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
Else
    Call SetWindowPos(Frm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End If
End Function

' Sub untuk menambahkan informasi ke Listview sampai 3 sub item
Public Sub AddInfoToList(lv As ucListView, sItem As String, sSub1 As String, sSub2 As String, sSub3 As String, iIcon As Long, nScroll As Long)
Dim lstLV As cListItem

Set lstLV = lv.ListItems.Add(, sItem, , iIcon, , , , , "")
    lstLV.SubItem(2).Text = sSub1
    lstLV.SubItem(3).Text = sSub2
    lstLV.SubItem(4).Text = sSub3
If lv.ListItems.count > nScroll Then lv.Scroll 0, 25

Set lstLV = Nothing
End Sub




