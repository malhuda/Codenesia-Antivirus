VERSION 5.00
Begin VB.Form frmRTP 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.C P.r.o.t.e.c.t.o.r"
   ClientHeight    =   3120
   ClientLeft      =   3735
   ClientTop       =   3885
   ClientWidth     =   8925
   ControlBox      =   0   'False
   Icon            =   "frmRTP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "frmRTP.frx":000C
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
   Begin VB.CommandButton cmdFixAllRtp 
      Caption         =   "FIX ALL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   3960
      TabIndex        =   3
      Top             =   2660
      Width           =   1575
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "Ignore "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   7200
      TabIndex        =   2
      Top             =   2660
      Width           =   1575
   End
   Begin VB.CommandButton cmdFixRtp 
      Caption         =   "FIX Selected"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   5640
      TabIndex        =   1
      Top             =   2660
      Width           =   1455
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   960
      TabIndex        =   0
      Top             =   2660
      Width           =   1575
   End
   Begin CMC.ucListView lvRTP 
      Height          =   2415
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4260
      StyleEx         =   33
      ShowSort        =   -1  'True
   End
End
Attribute VB_Name = "frmRTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFixAllRtp_Click()
    FiX_Malware lvRTP, BY_ALL, 10
End Sub

Private Sub cmdFixRtp_Click()
    FiX_Malware lvRTP, BY_SELECT, 10
End Sub

Private Sub cmdIgnore_Click()
    Call LepasSemuaKunci
    Me.Hide
    lvRTP.ListItems.Clear
End Sub

Private Sub cmdTutup_Click()
    Me.Hide
    lvRTP.ListItems.Clear
End Sub

Private Sub Form_Initialize()
 Call LetakanForm(Me, True)
End Sub


Private Sub Form_Load()
 Call LetakanForm(Me, True)
End Sub

Private Sub lvRTP_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub
