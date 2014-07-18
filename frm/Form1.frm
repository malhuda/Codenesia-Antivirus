VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "modDesain aj"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      Height          =   6495
      Index           =   0
      Left            =   0
      ScaleHeight     =   6495
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7935
      Begin CMC.uTabSonny TabMain 
         Height          =   6255
         Left            =   140
         TabIndex        =   1
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11033
         tabcount        =   5
         judul(1)        =   "Path Scan"
         judul(2)        =   "Malware"
         judul(3)        =   "Registry"
         judul(4)        =   "Hidden"
         judul(5)        =   "Information"
         Begin VB.PictureBox picPath 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            Height          =   5775
            Left            =   11950
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   21
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdStartScan 
               Caption         =   "Start Scan"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   0
               TabIndex        =   22
               Top             =   3840
               Width           =   1815
            End
            Begin CMC.ucProgressBar PB1 
               Height          =   255
               Left            =   0
               Top             =   5450
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   450
               Smooth          =   -1  'True
            End
            Begin CMC.DirTree DirTree 
               Height          =   3615
               Left            =   0
               TabIndex        =   23
               Top             =   120
               Width           =   7440
               _extentx        =   13123
               _extenty        =   6376
            End
            Begin VB.Label lbStatus 
               BackStyle       =   0  'Transparent
               Caption         =   " [Ready]"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   915
               TabIndex        =   43
               Top             =   4920
               Width           =   4095
            End
            Begin VB.Label lbMalware1 
               BackStyle       =   0  'Transparent
               Caption         =   "Malware "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5040
               TabIndex        =   42
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label lbRegistry1 
               BackStyle       =   0  'Transparent
               Caption         =   "Registry"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   255
               Left            =   5040
               TabIndex        =   41
               Top             =   4080
               Width           =   975
            End
            Begin VB.Label lbHidden1 
               BackStyle       =   0  'Transparent
               Caption         =   "Hidden"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   255
               Left            =   5040
               TabIndex        =   40
               Top             =   4320
               Width           =   975
            End
            Begin VB.Label lbMalware 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 000000 object(s)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   6000
               TabIndex        =   39
               Top             =   3840
               Width           =   1335
            End
            Begin VB.Label lbReg 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 000000 object(s)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   195
               Left            =   6000
               TabIndex        =   38
               Top             =   4080
               Width           =   1335
            End
            Begin VB.Label lbHidden 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 000000 object(s)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   195
               Left            =   6000
               TabIndex        =   37
               Top             =   4320
               Width           =   1335
            End
            Begin VB.Label lbInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 000000 object(s)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   6000
               TabIndex        =   36
               Top             =   4560
               Width           =   1335
            End
            Begin VB.Label lbObject 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   915
               TabIndex        =   35
               Top             =   5160
               Width           =   285
            End
            Begin VB.Label lblProcessed 
               BackStyle       =   0  'Transparent
               Caption         =   "Processed : "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   0
               TabIndex        =   34
               Top             =   5160
               Width           =   855
            End
            Begin VB.Label lbStatus1 
               BackStyle       =   0  'Transparent
               Caption         =   "Status       : "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   0
               TabIndex        =   33
               Top             =   4920
               Width           =   855
            End
            Begin VB.Label lbTime1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3000
               TabIndex        =   32
               Top             =   3840
               Width           =   330
            End
            Begin VB.Label lbFileFound1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Founded"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3000
               TabIndex        =   31
               Top             =   4080
               Width           =   630
            End
            Begin VB.Label lbFileCheck1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Checked"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3000
               TabIndex        =   30
               Top             =   4320
               Width           =   615
            End
            Begin VB.Label lbBypass1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ByPassed"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3000
               TabIndex        =   29
               Top             =   4560
               Width           =   690
            End
            Begin VB.Label lbTime 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 00 :00 :00"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3960
               TabIndex        =   28
               Top             =   3840
               Width           =   855
            End
            Begin VB.Label lbFileFound 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 00000000"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3960
               TabIndex        =   27
               Top             =   4080
               Width           =   825
            End
            Begin VB.Label lbFileCheck 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 00000000 "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3960
               TabIndex        =   26
               Top             =   4320
               Width           =   870
            End
            Begin VB.Label lbBypass 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ": 00000000"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3960
               TabIndex        =   25
               Top             =   4560
               Width           =   825
            End
            Begin VB.Label lblInfor 
               BackStyle       =   0  'Transparent
               Caption         =   "Information"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5040
               TabIndex        =   24
               Top             =   4560
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00F1DFB3&
            Height          =   5775
            Left            =   -23540
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   16
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdFixHiddenAll 
               Caption         =   "Fix All Object"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1800
               TabIndex        =   19
               Top             =   5280
               Width           =   1575
            End
            Begin VB.CommandButton cmdFixHidden 
               Caption         =   "Fix Checked"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   18
               Top             =   5280
               Width           =   1575
            End
            Begin VB.PictureBox picInfoHidden 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   6840
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   17
               Top             =   5280
               Width           =   495
            End
            Begin CMC.ucListView lvHidden 
               Height          =   5055
               Left            =   120
               TabIndex        =   20
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   37
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            Height          =   5775
            Left            =   -35370
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   11
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdProperties 
               Caption         =   "Properties"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4080
               TabIndex        =   14
               Top             =   5280
               Width           =   1575
            End
            Begin VB.CommandButton cmdExplore 
               Caption         =   "Explore"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5760
               TabIndex        =   13
               Top             =   5280
               Width           =   1575
            End
            Begin VB.PictureBox picIconInfo 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   120
               ScaleHeight     =   495
               ScaleWidth      =   615
               TabIndex        =   12
               Top             =   5280
               Width           =   615
            End
            Begin CMC.ucListView lvInfo 
               Height          =   5055
               Left            =   120
               TabIndex        =   15
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   33
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            Height          =   5775
            Left            =   -11710
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   7
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdFixReg 
               Caption         =   "Fix Checked"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   9
               Top             =   5280
               Width           =   1575
            End
            Begin VB.CommandButton cmdFixRegAll 
               Caption         =   "Fix All Object"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1800
               TabIndex        =   8
               Top             =   5280
               Width           =   1575
            End
            Begin CMC.ucListView lvRegistry 
               Height          =   5055
               Left            =   120
               TabIndex        =   10
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   37
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            Height          =   5775
            Index           =   0
            Left            =   120
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   2
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdFixMalwareAll 
               Caption         =   "Fix All Object"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1800
               TabIndex        =   5
               Top             =   5280
               Width           =   1575
            End
            Begin VB.CommandButton cmdFixMalware 
               Caption         =   "Fix Checked"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   4
               Top             =   5280
               Width           =   1575
            End
            Begin VB.PictureBox picInfoMalware 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   495
               Left            =   6840
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   33
               TabIndex        =   3
               Top             =   5280
               Width           =   495
            End
            Begin CMC.ucListView lvMalware 
               Height          =   5055
               Left            =   120
               TabIndex        =   6
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   37
               ShowSort        =   -1  'True
            End
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LngUpdate As Long

Private Sub cmdCheckUpdate_Click()

End Sub

Private Sub tmAnimUpdate_Timer()
Dim SKarUpdStat As String
LngUpdate = LngUpdate + 1

SKarUpdStat = Left(j_bahasa(26), LngUpdate)
If LngUpdate = Len(j_bahasa(26)) Then LngUpdate = 0

lblStatusUpdate.Caption = SKarUpdStat
End Sub
