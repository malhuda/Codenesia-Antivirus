Attribute VB_Name = "basRTP"
' Public Varnya disini sebagian
Public StatusRTP   As Boolean ' Klo False berarti RTP mati

Public Sub ScanPatWithRTP(PathWillScan As String)
' Cek apakah path termasuk Pengecualian
If ApaPengecualian(PathWillScan, JumPathExcep) = True Then GoTo BERAKHIR
If StatusRTP = False Then GoTo BERAKHIR
   
   DoEvents
   
   ScanRTP PathWillScan
If frmRTP.lvRTP.ListItems.count > 0 Then
With frmRTP
     .Show
     LetakanForm frmRTP, True
     .SetFocus
     .Caption = "C.M.C P.r.o.t.e.c.t.o.r [ " & .lvRTP.ListItems.count & " ]"
     KunciFileYangDiRTP .lvRTP
End With
End If
BERAKHIR:
End Sub


' Private karena untuk module ini saja
Private Function ApaPengecualian(spath As String, JumlahPath As Long) As Boolean
Dim iCount As Integer
On Error GoTo LBL_AKHIR
For iCount = 1 To JumlahPath
    If InStr(UCase(spath), UCase(PathExcep(iCount))) > 0 Then
       ApaPengecualian = True
       Exit Function
    End If
Next
LBL_AKHIR:
End Function


' Publik karena mau diakses dimana-mana
Public Function ApaPengecualianFile(sFile As String, JumFileExc As Long) As Boolean
Dim iCount As Integer

If sFile = "" Then Exit Function

For iCount = 1 To JumFileExc
    If UCase(FileExcep(iCount)) = UCase(sFile) Then
       ApaPengecualianFile = True
       Exit Function
    End If
Next

End Function

' Publik karena mau diakses dimana-mana (senearnya cuma di modReg dan DbReg aj)
Public Function ApaPengecualianReg(sRegPathAndValue As String, JumRegExc As Long) As Boolean
Dim iCount As Integer

If sRegPathAndValue = "" Then Exit Function

For iCount = 1 To JumRegExc
    If UCase(RegExcep(iCount)) = UCase(sRegPathAndValue) Then
       ApaPengecualianReg = True
       Exit Function
    End If
Next


End Function

' Yang masuk di RTP coba kunci semua
Public Sub KunciFileYangDiRTP(listRTP As ucListView)
Dim CountToBeLock   As Long
Dim PthKunci        As String
On Error Resume Next

For CountToBeLock = 1 To listRTP.ListItems.count
    PthKunci = listRTP.ListItems.Item(CountToBeLock).SubItem(2).Text
    KunciFile PthKunci
Next
End Sub

