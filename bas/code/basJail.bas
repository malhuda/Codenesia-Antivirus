Attribute VB_Name = "basJail"


Private cHufman As New classHuffman

'------------------------------------ Memenjarakan File ke Penjara
' Struktur file dipenjara
' ekstensi : n.$$
' CMS!?AlamatAsal?UkuranAsli?Virus?[Konst Pemisah]-Dataterenkrispi
' contoh CMS!C:\Aku.exe|100|Alman|*****|isienkirpsi
Public Function JailFile(sFile As String, sFolder As String, sVirName As String) As Long
Dim clFile      As New classFile
Dim nOP         As Long
Dim hFile       As Long
Dim nLenght     As Long
Dim nTurn       As Long
Dim sTmp        As String
Dim sHead       As String
Dim nFree       As String

'''KET
' 0 = gagal mempenjarakan dan hapus
' 1 = berhasil memenjarakan tapi gagal hapus
' 2 = berhasil semuanya

On Error GoTo LBL_GAGAL
BuatPenjara ' buat jaga-jaga aj
hFile = clFile.VbOpenFile(sFile, FOR_BINARY_ACCESS_READ, LOCK_NONE)
If hFile > 0 Then
   Dim BuffData()  As Byte
   nLenght = clFile.VbFileLen(hFile)
   nOP = clFile.VbReadFileB(hFile, 1, nLenght, BuffData)
   clFile.VbCloseFile hFile
   ' enkripsi cuy   100 data pertama aj
   For nTurn = 0 To 99
      BuffData(nTurn) = BuffData(nTurn) Xor 5
   Next
   ' Rapetin dulu
   cHufman.EncodeByte BuffData(), UBound(BuffData) + 1
   
   ' jadikan string cuy
   sTmp = StrConv(BuffData, vbUnicode)
   
   sHead = "CMS!" & sFile & "|" & nLenght & "|" & sVirName & "|"
   
   sTmp = StrConv(sHead, vbUnicode) & "*****|" & sTmp 'klo sFile disi char unicode hasilnya kok masih "?" yah :(
   nFree = GetFreeKar(sFolder)
   WriteFileUniSim sFolder & "\" & nFree & JailExt, sTmp
   
   HapusFile sFile
   
   If ValidFile(sFile) = True Then ' gagal hapus
      JailFile = 1
   Else
      JailFile = 2
   End If

End If

Exit Function
LBL_GAGAL:
JailFile = 0
End Function

Public Function RestoreJailFile(sFile As String, sTarget As String)
Dim clFile      As New classFile
Dim nFree       As Long
Dim nOP         As Long
Dim hFile       As Long
Dim nLenght     As Long
Dim nTurn       As Long
Dim sTmp        As String
Dim nPos        As Long
On Error Resume Next

hFile = clFile.VbOpenFile(sFile, FOR_BINARY_ACCESS_READ, LOCK_NONE)

If hFile > 0 Then
 ' deteksi dulu sparatornya
   Dim BuffData()  As Byte
   nLenght = clFile.VbFileLen(hFile)
   If nLenght > 1000 Then nPos = 1000 Else nPos = nLenght '--- baca 1000 dari depan aj
   nOP = clFile.VbReadFileB(hFile, 1, nPos, BuffData)
   sTmp = StrConv(BuffData, vbUnicode)
   
   nPos = InStr(sTmp, "*****|") + 6
   
   Erase BuffData ' hapus dulu mau dipakai

   nLenght = nLenght - nPos
   nOP = clFile.VbReadFileB(hFile, nPos, nLenght + 1, BuffData)
   
   sTmp = StrConv(BuffData, vbUnicode)
   clFile.VbCloseFile hFile
   ' Melarin dulu
   cHufman.DecodeByte BuffData(), UBound(BuffData) + 1
      
   ' enkripsi cuy   100 data pertama aj
   For nTurn = 0 To 99
      BuffData(nTurn) = BuffData(nTurn) Xor 5
   Next
   
   ' jadikan string cuy
   sTmp = StrConv(BuffData, vbUnicode)
   WriteFileUniSim sTarget, sTmp
End If
End Function

Private Function GetFreeKar(sFolder As String) As String
Dim nFile As Long
nFile = 256 ' langsung keluar dari UNI
Do
    If ValidFile(sFolder & "\" & ChrW$(nFile) & JailExt) = False Then GoSub LBL_AKHIR
    nFile = nFile + 1
Loop

Exit Function
LBL_AKHIR:
GetFreeKar = ChrW$(nFile)
End Function

Public Sub BuatPenjara()
Dim isinya As String
If PathIsDirectory(StrPtr(FolderJail)) = 0 Then
   BuatFolder FolderJail
   SetFileAttributes StrPtr(FolderJail), vbHidden
   isinya = " - Codenesia Malware Cleaner - " & vbNewLine & vbNewLine
   isinya = isinya & "Folder ini tempat penjara para tahanan C.M.C ! " & ChrW(&H263A) & vbNewLine & vbNewLine
   isinya = isinya & ChrW(&H2665) & " CM Team - Made in Indonesia " & ChrW(&H2665)
   isinya = StrConv(isinya, vbUnicode)
   WriteFileUniSim FolderJail & "\R" & ChrW$(&H20AC) & ChrW$(&H20AA) & "d.txt", isinya
Else
   SetFileAttributes StrPtr(FolderJail), vbHidden
End If
End Sub


' Membaca Data pada folder penjara
Public Function READ_DATA_JAIL(sJail As String)
Dim nFree       As Long
Dim lngItem     As Long
Dim nOP         As Long
Dim hFile       As Long
Dim nLenght     As Long
Dim nTurn       As Long
Dim spath       As String
Dim sTmp        As String
Dim nVirName    As String
Dim aTmp()      As String
Dim sFile()     As String
Dim nFile       As Long
Dim clFile As New classFile


nFile = GetFile(sJail, sFile)
With frmMain.lvJail
     .ListItems.Clear

For nTurn = 0 To nFile - 1
    On Error GoTo LEWAT
    hFile = clFile.VbOpenFile(sFile(nTurn), FOR_BINARY_ACCESS_READ, LOCK_NONE)
    If hFile > 0 Then
        Dim BuffData()  As Byte
        nLenght = clFile.VbFileLen(hFile)
        If nLenght > 1000 Then nLenght = 1000
        nOP = clFile.VbReadFileB(hFile, 1, nLenght, BuffData) ' baca 100 data terdepan aj
              clFile.VbCloseFile hFile ' lgsung tutup
        sTmp = StrConv(BuffData, vbUnicode)
        sTmp = StrConv(sTmp, vbFromUnicode)
        aTmp() = Split(sTmp, "|")
        spath = Mid(aTmp(0), 5)
        nLenght = CLng(aTmp(1))
        nVirName = aTmp(2)
        .ListItems.Add , nVirName, , 0, , , , , Array(spath, Format$(nLenght, "#,#"), GetFileName(sFile(nTurn)))
        Erase BuffData
    End If
LEWAT:
Next

frmMain.frJail = c_bahasa(6) & " ( " & .ListItems.Count & " )"
End With
End Function


'------------------- Command2 Function

Public Sub ClearJail(lvJail As ucListView)
Dim nCount As Long
Dim sFile  As String
On Error Resume Next
With lvJail
If .ListItems.Count = 0 Then Exit Sub

For nCount = 0 To .ListItems.Count - 1
    sFile = FolderJail & "\" & .ListItems.Item(nCount + 1).SubItem(4).Text
    HapusFile sFile
Next

lvJail.ListItems.Clear

MsgBox i_bahasa(0), vbExclamation

Call READ_DATA_JAIL(FolderJail)
End With
End Sub

Public Sub KillPrisonner(PrisonerName As String, lvJail As ucListView)
Dim sFile As String

On Error Resume Next

With lvJail
If .ListItems.Count = 0 Then Exit Sub
    sFile = FolderJail & "\" & PrisonerName
    HapusFile sFile
End With
End Sub

Public Sub ReleasePrisoner(OriPathFile As String, sFileNamePrisoner, lvJail As ucListView)
Dim sFile   As String
With lvJail
    If .ListItems.Count = 0 Then Exit Sub
    sFile = FolderJail & "\" & sFileNamePrisoner
    If PathIsDirectory(StrPtr(GetFilePath(OriPathFile))) <> 0 Then
       If ValidFile(OriPathFile) = True Then
          If MsgBox(i_bahasa(1), vbYesNo + vbExclamation) = vbYes Then
             RestoreJailFile sFile, OriPathFile
             HapusFile sFile
             MsgBoxU i_bahasa(2) & " ( " & GetFilePath(OriPathFile) & " )", App.Title, vbExclamation, frmMain
             READ_DATA_JAIL FolderJail
          End If
       Else
            RestoreJailFile sFile, OriPathFile
            HapusFile sFile
            MsgBoxU i_bahasa(2) & " ( " & GetFilePath(OriPathFile) & " )", App.Title, vbExclamation, frmMain
            READ_DATA_JAIL FolderJail
       End If
     Else
        MsgBoxU i_bahasa(3), App.Title, vbExclamation, frmMain
     End If
End With
End Sub


