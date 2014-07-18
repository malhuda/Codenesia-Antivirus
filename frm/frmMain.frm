VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "C.M.C  PH.3-5 - Final Version"
   ClientHeight    =   7440
   ClientLeft      =   2595
   ClientTop       =   2175
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10710
   Begin VB.Timer tmFlash 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   2040
      Top             =   6480
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   5
      Left            =   2400
      ScaleHeight     =   6015
      ScaleWidth      =   7935
      TabIndex        =   182
      Top             =   1680
      Visible         =   0   'False
      Width           =   7935
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillColor       =   &H0080FF80&
         Height          =   6255
         Index           =   1
         Left            =   140
         ScaleHeight     =   6255
         ScaleWidth      =   7680
         TabIndex        =   183
         Top             =   120
         Width           =   7680
         Begin VB.PictureBox picPlugin 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7080
            ScaleHeight     =   33
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   33
            TabIndex        =   222
            Top             =   4680
            Width           =   495
         End
         Begin VB.CommandButton cmdExecutePlug 
            Caption         =   "Execute Plugin"
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
            Left            =   5880
            TabIndex        =   185
            Top             =   4200
            Width           =   1695
         End
         Begin VB.ListBox lstPlugin 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3765
            Left            =   120
            TabIndex        =   184
            Top             =   360
            Width           =   7455
         End
         Begin VB.Label lblPlugVer 
            Height          =   255
            Left            =   120
            TabIndex        =   223
            Top             =   5160
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblPlugDesc1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                                  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1560
            TabIndex        =   197
            Top             =   5520
            Width           =   5895
         End
         Begin VB.Label lblPlugAut1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   196
            Top             =   4440
            Width           =   960
         End
         Begin VB.Label lblPlugSelect1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   195
            Top             =   4200
            Width           =   960
         End
         Begin VB.Label lblPlugDesc 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   194
            Top             =   5520
            Width           =   1260
         End
         Begin VB.Label lblPlugAut 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Author"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   193
            Top             =   4440
            Width           =   960
         End
         Begin VB.Label lblPlugSelect 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Plugin Selected"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   192
            Top             =   4200
            Width           =   1080
         End
         Begin VB.Label lblAvalaiblePlug 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Avalaible Plugin(s)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   191
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label lblPlugAutEmail 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Author Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   190
            Top             =   4680
            Width           =   900
         End
         Begin VB.Label lblPlugAutEmail1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   189
            Top             =   4680
            Width           =   2280
         End
         Begin VB.Label lblPlugAutSite 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Author Site"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   188
            Top             =   4920
            Width           =   810
         End
         Begin VB.Label lblPlugAutSite1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1560
            TabIndex        =   187
            Top             =   4920
            Width           =   2280
         End
         Begin VB.Label lblPlugVer1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   ":                    "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7080
            TabIndex        =   186
            Top             =   4800
            Visible         =   0   'False
            Width           =   2280
         End
      End
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      Height          =   5415
      Index           =   4
      Left            =   3240
      ScaleHeight     =   5415
      ScaleWidth      =   7935
      TabIndex        =   127
      Top             =   1440
      Visible         =   0   'False
      Width           =   7935
      Begin CMC.uTabSonny TabAbout 
         Height          =   6255
         Left            =   140
         TabIndex        =   128
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11033
         tabcount        =   2
         judul(1)        =   "About CMC"
         judul(2)        =   "C.M.C Information"
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   5775
            Index           =   1
            Left            =   120
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   150
            Top             =   360
            Width           =   7455
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PmZ"
               Height          =   210
               Index           =   33
               Left            =   360
               TabIndex        =   221
               Top             =   3600
               Width           =   315
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bidan Malware"
               Height          =   210
               Index           =   32
               Left            =   360
               TabIndex        =   220
               Top             =   3360
               Width           =   1080
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   "Virus Analyst"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   31
               Left            =   360
               TabIndex        =   219
               Top             =   3120
               Width           =   1575
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Terren.Jr"
               Height          =   210
               Index           =   30
               Left            =   6240
               TabIndex        =   218
               Top             =   1680
               Width           =   660
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "cmc@codenesia.com"
               Height          =   210
               Index           =   29
               Left            =   5280
               TabIndex        =   198
               Top             =   3600
               Width           =   1560
            End
            Begin VB.Label lblAbout 
               BackColor       =   &H00F1DFB3&
               BackStyle       =   0  'Transparent
               Caption         =   "Codenesia Malware Team"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   1200
               TabIndex        =   179
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label lblAbout 
               BackColor       =   &H00F1DFB3&
               BackStyle       =   0  'Transparent
               Caption         =   "Codenesia Malware Cleaner"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   960
               TabIndex        =   178
               Top             =   120
               Width           =   4815
            End
            Begin VB.Image Image1 
               Height          =   960
               Index           =   1
               Left            =   0
               Picture         =   "frmMain.frx":5BBA
               Top             =   0
               Width           =   960
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   "Code Developer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   177
               Top             =   1200
               Width           =   1935
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HrXxX"
               Height          =   210
               Index           =   5
               Left            =   360
               TabIndex        =   176
               Top             =   1440
               Width           =   465
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "APTX"
               Height          =   210
               Index           =   6
               Left            =   360
               TabIndex        =   175
               Top             =   1680
               Width           =   405
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   "Malware Hunter"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   174
               Top             =   2280
               Width           =   1575
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gxry"
               Height          =   210
               Index           =   8
               Left            =   360
               TabIndex        =   173
               Top             =   2520
               Width           =   360
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Revil"
               Height          =   210
               Index           =   9
               Left            =   1200
               TabIndex        =   172
               Top             =   2520
               Width           =   345
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Heru"
               Height          =   210
               Index           =   10
               Left            =   1200
               TabIndex        =   171
               Top             =   2760
               Width           =   345
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cheopz"
               Height          =   210
               Index           =   11
               Left            =   5280
               TabIndex        =   170
               Top             =   1920
               Width           =   555
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Anharku"
               Height          =   210
               Index           =   12
               Left            =   360
               TabIndex        =   169
               Top             =   2760
               Width           =   615
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sonny"
               Height          =   210
               Index           =   16
               Left            =   2520
               TabIndex        =   168
               Top             =   2520
               Width           =   465
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Program Designer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   15
               Left            =   2520
               TabIndex        =   167
               Top             =   2280
               Width           =   1515
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Code Reviewer n  Advicer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   2520
               TabIndex        =   166
               Top             =   1200
               Width           =   2130
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pamzlogic"
               Height          =   210
               Index           =   4
               Left            =   2520
               TabIndex        =   165
               Top             =   1440
               Width           =   720
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   "Support Publisher"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   5280
               TabIndex        =   164
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "R-Win"
               Height          =   210
               Index           =   23
               Left            =   5280
               TabIndex        =   163
               Top             =   1680
               Width           =   435
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VirSpector"
               Height          =   210
               Index           =   22
               Left            =   5280
               TabIndex        =   162
               Top             =   1440
               Width           =   780
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kholis"
               Height          =   210
               Index           =   13
               Left            =   6240
               TabIndex        =   161
               Top             =   1440
               Width           =   435
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   "Web Developer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   2520
               TabIndex        =   160
               Top             =   3120
               Width           =   1575
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ajrnea"
               Height          =   210
               Index           =   18
               Left            =   2520
               TabIndex        =   159
               Top             =   3360
               Width           =   480
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Situs  Kami"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   19
               Left            =   5280
               TabIndex        =   158
               Top             =   2280
               Width           =   915
            End
            Begin VB.Label lblAbout 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C.M.C ## Copyright@2009-2010 - CodenesiaSoft"
               Height          =   210
               Index           =   24
               Left            =   3915
               TabIndex        =   157
               Top             =   5520
               Width           =   3540
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "www.codenesia.com"
               Height          =   210
               Index           =   20
               Left            =   5280
               TabIndex        =   156
               Top             =   2520
               Width           =   1590
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "www.cmc.codenesia.com"
               Height          =   210
               Index           =   21
               Left            =   5280
               TabIndex        =   155
               Top             =   2760
               Width           =   1935
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   25
               Left            =   5280
               TabIndex        =   154
               Top             =   3120
               Width           =   435
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "im4soft@gmail.com"
               Height          =   210
               Index           =   26
               Left            =   5280
               TabIndex        =   153
               Top             =   3360
               Width           =   1395
            End
            Begin VB.Label lblAbout 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Adsense"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   27
               Left            =   360
               TabIndex        =   152
               Top             =   4080
               Width           =   750
            End
            Begin VB.Label lblAbout 
               BackStyle       =   0  'Transparent
               Caption         =   $"frmMain.frx":B774
               Height          =   855
               Index           =   28
               Left            =   360
               TabIndex        =   151
               Top             =   4320
               Width           =   6135
            End
         End
         Begin VB.PictureBox PicCmcInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   5775
            Left            =   -11710
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   129
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdMoreInfo 
               Caption         =   "Click For More Information"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4200
               TabIndex        =   149
               Top             =   5280
               Width           =   3255
            End
            Begin VB.Frame frSoftInformation 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Software Information"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   4200
               TabIndex        =   134
               Top             =   0
               Width           =   3255
               Begin VB.Label lbMesin 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 32 bit - Windows XP"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   148
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Machine"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   147
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.Label lbVirus 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 0009 + Heuristic"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   146
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.Label lbWorm 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": -"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   145
                  Top             =   1200
                  Width           =   1695
               End
               Begin VB.Label lbRegDataBase 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 106 value(s)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   144
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.Label lbBuildDate 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 16 Maret 2010"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   143
                  Top             =   720
                  Width           =   1695
               End
               Begin VB.Label lbBuildNumber 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": 5"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   142
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.Label lbEngine 
                  BackStyle       =   0  'Transparent
                  Caption         =   ": PH#3"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   141
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Internal Virus"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   140
                  Top             =   1440
                  Width           =   1215
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Internal Worm"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   139
                  Top             =   1200
                  Width           =   1215
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Reg Database"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   138
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Build Date"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   137
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Build Number"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   136
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.Label lbInfo1 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Engine Version "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   135
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.Frame frInteralMalware 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Internal Malware List"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5775
               Left            =   0
               TabIndex        =   132
               Top             =   0
               Width           =   4095
               Begin VB.ListBox lstListWorm 
                  Height          =   5310
                  Left            =   120
                  Sorted          =   -1  'True
                  TabIndex        =   133
                  Top             =   240
                  Width           =   3855
               End
            End
            Begin VB.Frame frInternalVirus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Internal Virus Detector"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3015
               Left            =   4200
               TabIndex        =   130
               Top             =   2160
               Width           =   3255
               Begin VB.ListBox lstVirus 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2595
                  ItemData        =   "frmMain.frx":B888
                  Left            =   120
                  List            =   "frmMain.frx":B8A7
                  TabIndex        =   131
                  Top             =   240
                  Width           =   3015
               End
            End
         End
      End
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Index           =   3
      Left            =   1560
      ScaleHeight     =   2295
      ScaleWidth      =   7935
      TabIndex        =   121
      Top             =   3120
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton cmdCheckUpdate 
         Caption         =   " Update Now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   123
         Top             =   5160
         Width           =   2295
      End
      Begin VB.TextBox txtRetriveInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   122
         Top             =   120
         Width           =   7695
      End
      Begin CMC.ucProgressBar PB_UPD 
         Height          =   250
         Left            =   120
         Top             =   4080
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   450
      End
      Begin CMC.ucProgressBar PBC 
         Height          =   255
         Left            =   120
         Top             =   4440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Smooth          =   -1  'True
      End
      Begin VB.Label lblStatusUpdate 
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
         Height          =   255
         Left            =   140
         TabIndex        =   124
         Top             =   4800
         Width           =   7695
      End
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      ScaleHeight     =   375
      ScaleWidth      =   7935
      TabIndex        =   78
      Top             =   4680
      Visible         =   0   'False
      Width           =   7935
      Begin CMC.uTabSonny TabConfig 
         Height          =   6255
         Left            =   140
         TabIndex        =   79
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11033
         tabcount        =   5
         judul(1)        =   "Application"
         judul(2)        =   "Language"
         judul(3)        =   "RTP Exception(s)"
         judul(4)        =   "File Exception(s)"
         judul(5)        =   "Registry Exception(s)"
         Begin VB.PictureBox PicConfigApp 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   120
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   101
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save Configuration"
               Height          =   495
               Left            =   5640
               TabIndex        =   102
               Top             =   5160
               Width           =   1695
            End
            Begin CMC.ucFrame FrConfigScan 
               Height          =   2415
               Left            =   120
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   4260
               Caption         =   "Scan Option"
               BColor          =   16777215
               Begin VB.CheckBox ck6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable unpack archive (zip, rar. gz, tgz)"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   108
                  Top             =   2040
                  Width           =   5535
               End
               Begin VB.CheckBox ck1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable filter file (by pass file with certain extensions) "
                  Height          =   255
                  Left            =   120
                  TabIndex        =   107
                  ToolTipText     =   "Make slower but total scanning if unchecked"
                  Top             =   240
                  Width           =   4815
               End
               Begin VB.CheckBox ck2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable use Heuristic to suspect malware"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   106
                  ToolTipText     =   "Use C*M*S heuristic to suspect some malwares !"
                  Top             =   600
                  Value           =   1  'Checked
                  Width           =   4815
               End
               Begin VB.CheckBox ck4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable detect hidden object (file and folder)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   105
                  ToolTipText     =   "Make speed slower but detect hidden file and folder !"
                  Top             =   1320
                  Value           =   1  'Checked
                  Width           =   5535
               End
               Begin VB.CheckBox ck5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable give strange  information while scanning"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   104
                  Top             =   1680
                  Value           =   1  'Checked
                  Width           =   6015
               End
               Begin VB.CheckBox ck3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable detect useless registry value (XP only)"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   103
                  ToolTipText     =   "Only for XP OS"
                  Top             =   960
                  Value           =   1  'Checked
                  Width           =   5295
               End
            End
            Begin CMC.ucFrame FrConfigApp 
               Height          =   2415
               Left            =   120
               Top             =   2640
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   4260
               Caption         =   "Application Configuration"
               BColor          =   16777215
               Begin VB.CheckBox ck12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Install Context Menu -Scan With CMC-"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   180
                  Top             =   2040
                  Width           =   4935
               End
               Begin VB.CheckBox ck8 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable CMC Protection"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   113
                  Top             =   600
                  Width           =   4215
               End
               Begin VB.CheckBox ck7 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Enable Run on Startup"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   112
                  Top             =   240
                  Width           =   4215
               End
               Begin VB.CheckBox ck9 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Auto Check Online Update"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   111
                  Top             =   960
                  Width           =   4215
               End
               Begin VB.CheckBox ck10 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Auto Scan Flashdisk inserted"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   110
                  Top             =   1320
                  Width           =   4215
               End
               Begin VB.CheckBox ck11 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Place Application on Top"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   109
                  Top             =   1680
                  Width           =   4215
               End
            End
         End
         Begin VB.PictureBox PicConfLang 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -11710
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   94
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdApplyLang 
               Caption         =   "Apply Language"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5760
               TabIndex        =   96
               Top             =   4200
               Width           =   1575
            End
            Begin VB.ListBox lstLanguage 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3570
               Left            =   120
               TabIndex        =   95
               Top             =   480
               Width           =   7215
            End
            Begin VB.Label lblLangUsed1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   1920
               TabIndex        =   126
               Top             =   5280
               Width           =   1395
            End
            Begin VB.Label lblLangUsed 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Language Used"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   120
               TabIndex        =   125
               Top             =   5280
               Width           =   1290
            End
            Begin VB.Label lblLangAut1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1920
               TabIndex        =   117
               Top             =   4680
               Width           =   1275
            End
            Begin VB.Label lblLangID1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1920
               TabIndex        =   116
               Top             =   4440
               Width           =   1515
            End
            Begin VB.Label lblLangSel1 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1920
               TabIndex        =   115
               Top             =   4200
               Width           =   1605
            End
            Begin VB.Label lblLangAut 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Language Author "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   100
               Top             =   4680
               Width           =   1290
            End
            Begin VB.Label lblLangID 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Language ID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   99
               Top             =   4440
               Width           =   915
            End
            Begin VB.Label lblLangSel 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Language Selected"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   98
               Top             =   4200
               Width           =   1365
            End
            Begin VB.Label lblAvalaibleLang 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Avalaible Language"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   97
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.PictureBox PicExcPath 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -23540
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   89
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdRemovePath1 
               Caption         =   "Remove Selected"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   120
               Top             =   5280
               Width           =   1455
            End
            Begin VB.ListBox lstExceptFolder 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4740
               Left            =   120
               TabIndex        =   92
               Top             =   480
               Width           =   7215
            End
            Begin VB.CommandButton cmdAddExcFolder 
               Caption         =   "Add Path"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   91
               Top             =   5280
               Width           =   1455
            End
            Begin VB.CommandButton cmdRemovePath 
               Caption         =   "Remove All"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   90
               Top             =   5280
               Width           =   1455
            End
            Begin VB.Label lblExceptFolder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "RTP Exception - Dont Give me Warning about Threat in this Path"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   120
               Width           =   7335
            End
         End
         Begin VB.PictureBox PicExcFile 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -35370
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   84
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdRemExcFile1 
               Caption         =   "Remove Selected"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1680
               TabIndex        =   119
               Top             =   5280
               Width           =   1455
            End
            Begin VB.ListBox lstExceptFile 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4740
               Left            =   120
               TabIndex        =   87
               Top             =   480
               Width           =   7215
            End
            Begin VB.CommandButton cmdAddExcFile 
               Caption         =   "Add File"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   86
               Top             =   5280
               Width           =   1455
            End
            Begin VB.CommandButton cmdRemExcFile 
               Caption         =   "Remove All"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   85
               Top             =   5280
               Width           =   1455
            End
            Begin VB.Label lblExceptFile 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "File Exception - I'am sure this is normal file, dont catch as a malware file(s) below"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   88
               Top             =   120
               Width           =   5805
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Index           =   2
            Left            =   -47200
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   80
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdRemExcReg1 
               Caption         =   "Remove Selected"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4200
               TabIndex        =   118
               Top             =   5280
               Width           =   1575
            End
            Begin VB.ListBox lstExceptReg 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4740
               Left            =   120
               TabIndex        =   82
               Top             =   480
               Width           =   7215
            End
            Begin VB.CommandButton cmdRemExcReg 
               Caption         =   "Remove All"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   81
               Top             =   5280
               Width           =   1455
            End
            Begin VB.Label lblExceptReg 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Registry Exception - I'am sure this is normal value. Don't catch as a bad value, value(s) below"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   120
               Width           =   6735
            End
         End
      End
   End
   Begin VB.Timer tmAwal 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   6600
   End
   Begin CMC.rtp_mode rtp_mode1 
      Index           =   0
      Left            =   5160
      Top             =   0
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Index           =   2
      Left            =   2520
      ScaleHeight     =   2775
      ScaleWidth      =   6615
      TabIndex        =   51
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin CMC.uTabSonny TabTool 
         Height          =   6255
         Left            =   140
         TabIndex        =   52
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   11033
         tabcount        =   3
         judul(1)        =   "Prosess Manager"
         judul(2)        =   "Temporary Malware Signer"
         judul(3)        =   "Jail Controller"
         Begin VB.PictureBox picTempMalware 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -11710
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   66
            Top             =   360
            Width           =   7455
            Begin VB.Frame frTemp 
               BackColor       =   &H00FFFFFF&
               Caption         =   "List of Temporary  (0)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4095
               Left            =   0
               TabIndex        =   74
               Top             =   1680
               Width           =   7455
               Begin VB.ListBox lstVirTemp 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3765
                  Left            =   120
                  TabIndex        =   75
                  Top             =   240
                  Width           =   7215
               End
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Cancel"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1440
               TabIndex        =   73
               Top             =   1200
               Width           =   1335
            End
            Begin VB.CommandButton cmdAddVirus 
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   72
               Top             =   1200
               Width           =   1335
            End
            Begin VB.Frame frVirus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Malware Temporary"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   7455
               Begin VB.CommandButton cmdBrowse 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   6960
                  TabIndex        =   76
                  Top             =   240
                  Width           =   395
               End
               Begin VB.TextBox txtVirusName 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1440
                  TabIndex        =   69
                  Top             =   720
                  Width           =   2295
               End
               Begin VB.TextBox txtVirusPath 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   68
                  Top             =   280
                  Width           =   5415
               End
               Begin VB.Label lblMalwareName 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Malware Name"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   71
                  Top             =   720
                  Width           =   1050
               End
               Begin VB.Label lblMalwarePath 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Malware Path"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   70
                  Top             =   285
                  Width           =   975
               End
            End
         End
         Begin VB.PictureBox PicProsesManager 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   120
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   60
            Top             =   360
            Width           =   7455
            Begin VB.Frame frProses 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Process List"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4095
               Left            =   0
               TabIndex        =   63
               Top             =   0
               Width           =   7455
               Begin CMC.ucListView lvProses 
                  Height          =   3735
                  Left            =   120
                  TabIndex        =   64
                  Top             =   240
                  Width           =   7215
                  _ExtentX        =   12726
                  _ExtentY        =   6588
                  StyleEx         =   33
               End
               Begin VB.Label lblSelectedPID 
                  Caption         =   "lblSelectedPID"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   135
                  Left            =   5280
                  TabIndex        =   65
                  Top             =   3960
                  Visible         =   0   'False
                  Width           =   1095
               End
            End
            Begin VB.Frame frModule 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Module List"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1575
               Left            =   0
               TabIndex        =   61
               Top             =   4200
               Width           =   7455
               Begin VB.ListBox lstModule 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1230
                  Left            =   120
                  TabIndex        =   62
                  Top             =   240
                  Width           =   7215
               End
            End
         End
         Begin VB.PictureBox PicJailCont 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -23540
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   53
            Top             =   360
            Width           =   7455
            Begin VB.CommandButton cmdClearJail 
               Caption         =   "Clear Jail"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   59
               Top             =   5400
               Width           =   1455
            End
            Begin VB.CommandButton cmdReleaseTo 
               Caption         =   "Release To.."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   4680
               TabIndex        =   58
               Top             =   5400
               Width           =   1575
            End
            Begin VB.CommandButton cmdKillPris 
               Caption         =   "Kill Prisoner"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   57
               Top             =   5400
               Width           =   1455
            End
            Begin VB.CommandButton cmdRelease 
               Caption         =   "Release.."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3120
               TabIndex        =   56
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Frame frJail 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Prisoner (0)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   5295
               Left            =   20
               TabIndex        =   54
               Top             =   0
               Width           =   7455
               Begin CMC.ucListView lvJail 
                  Height          =   4935
                  Left            =   120
                  TabIndex        =   55
                  Top             =   240
                  Width           =   7215
                  _ExtentX        =   12726
                  _ExtentY        =   8705
                  StyleEx         =   33
               End
            End
         End
      End
   End
   Begin VB.Timer tmTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   6840
   End
   Begin VB.PictureBox picTmpIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   39
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00F1DFB3&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Index           =   0
      Left            =   2880
      ScaleHeight     =   6495
      ScaleWidth      =   7935
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   7935
      Begin CMC.uTabSonny TabMain 
         Height          =   6255
         Left            =   140
         TabIndex        =   22
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
         Begin VB.PictureBox Picture5 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -47200
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   213
            Top             =   360
            Width           =   7455
            Begin VB.PictureBox picIconInfo 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               ScaleHeight     =   495
               ScaleWidth      =   615
               TabIndex        =   216
               Top             =   5280
               Width           =   615
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
               TabIndex        =   215
               Top             =   5280
               Width           =   1575
            End
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
               TabIndex        =   214
               Top             =   5280
               Width           =   1575
            End
            Begin CMC.ucListView lvInfo 
               Height          =   5055
               Left            =   120
               TabIndex        =   217
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   33
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H00F1DFB3&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -35370
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   208
            Top             =   360
            Width           =   7455
            Begin VB.PictureBox picInfoHidden 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6840
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   211
               Top             =   5280
               Width           =   495
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
               TabIndex        =   210
               Top             =   5280
               Width           =   1575
            End
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
               TabIndex        =   209
               Top             =   5280
               Width           =   1575
            End
            Begin CMC.ucListView lvHidden 
               Height          =   5055
               Left            =   120
               TabIndex        =   212
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   37
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   -23540
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   204
            Top             =   360
            Width           =   7455
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
               TabIndex        =   206
               Top             =   5280
               Width           =   1575
            End
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
               TabIndex        =   205
               Top             =   5280
               Width           =   1575
            End
            Begin CMC.ucListView lvRegistry 
               Height          =   5055
               Left            =   120
               TabIndex        =   207
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Index           =   0
            Left            =   -11710
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   199
            Top             =   360
            Width           =   7455
            Begin VB.PictureBox picInfoMalware 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   6840
               ScaleHeight     =   33
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   33
               TabIndex        =   202
               Top             =   5280
               Width           =   495
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
               TabIndex        =   201
               Top             =   5280
               Width           =   1575
            End
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
               TabIndex        =   200
               Top             =   5280
               Width           =   1575
            End
            Begin CMC.ucListView lvMalware 
               Height          =   5055
               Left            =   120
               TabIndex        =   203
               Top             =   120
               Width           =   7215
               _ExtentX        =   12726
               _ExtentY        =   8916
               StyleEx         =   37
               ShowSort        =   -1  'True
            End
         End
         Begin VB.PictureBox picPath 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillColor       =   &H0080FF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5775
            Left            =   120
            ScaleHeight     =   5775
            ScaleWidth      =   7455
            TabIndex        =   23
            Top             =   360
            Width           =   7455
            Begin CMC.ucProgressBar PB1 
               Height          =   255
               Left            =   0
               Top             =   5450
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   450
               Smooth          =   -1  'True
            End
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
               TabIndex        =   25
               Top             =   3840
               Width           =   1815
            End
            Begin CMC.DirTree DirTree 
               Height          =   3615
               Left            =   0
               TabIndex        =   24
               Top             =   120
               Width           =   7440
               _extentx        =   13123
               _extenty        =   6376
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
               TabIndex        =   114
               Top             =   4560
               Width           =   975
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
               TabIndex        =   48
               Top             =   4560
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
               TabIndex        =   47
               Top             =   4320
               Width           =   870
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
               TabIndex        =   46
               Top             =   4080
               Width           =   825
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
               TabIndex        =   45
               Top             =   3840
               Width           =   855
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
               TabIndex        =   44
               Top             =   4560
               Width           =   690
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
               TabIndex        =   43
               Top             =   4320
               Width           =   615
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
               TabIndex        =   42
               Top             =   4080
               Width           =   630
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
               TabIndex        =   41
               Top             =   3840
               Width           =   330
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
               TabIndex        =   40
               Top             =   4920
               Width           =   855
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
               TabIndex        =   38
               Top             =   5160
               Width           =   855
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
               TabIndex        =   34
               Top             =   5160
               Width           =   285
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
               TabIndex        =   33
               Top             =   4560
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
               TabIndex        =   32
               Top             =   4320
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
               TabIndex        =   31
               Top             =   4080
               Width           =   1335
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
               TabIndex        =   30
               Top             =   3840
               Width           =   1335
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
               TabIndex        =   29
               Top             =   4320
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
               TabIndex        =   28
               Top             =   4080
               Width           =   975
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
               TabIndex        =   27
               Top             =   3840
               Width           =   975
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
               TabIndex        =   26
               Top             =   4920
               Width           =   4095
            End
         End
      End
   End
   Begin VB.PictureBox PenggantiListImage1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   1275
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
      Begin VB.PictureBox picCaution 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Picture         =   "frmMain.frx":B9E7
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   77
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picFileHid 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Picture         =   "frmMain.frx":BD29
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   50
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox picFolHid 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Picture         =   "frmMain.frx":C06B
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   49
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox pic14 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Picture         =   "frmMain.frx":C3AD
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Picture         =   "frmMain.frx":C6C7
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   20
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic3 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Picture         =   "frmMain.frx":CA09
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pic4 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Picture         =   "frmMain.frx":CDBF
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic6 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Picture         =   "frmMain.frx":D101
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic7 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Picture         =   "frmMain.frx":D443
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic8 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Picture         =   "frmMain.frx":D9CD
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   15
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic9 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         Picture         =   "frmMain.frx":DF57
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic10 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Picture         =   "frmMain.frx":E4E1
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic11 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Picture         =   "frmMain.frx":EA6B
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic12 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Picture         =   "frmMain.frx":EFF5
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic5 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Picture         =   "frmMain.frx":F57F
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox pic13 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         Picture         =   "frmMain.frx":F8C1
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Picture         =   "frmMain.frx":FE4B
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":1018D
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   6
      Top             =   0
      Width           =   10815
      Begin VB.Timer tmUpdate 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   8280
         Top             =   360
      End
      Begin VB.Label lblWeb 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "www.codenesia.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   8040
         TabIndex        =   37
         ToolTipText     =   "Goto www.codenesia.com"
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      ScaleHeight     =   5055
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   960
      Width           =   2775
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Scan"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":15FC7
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Configuration"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":18E9B
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         MaskColor       =   16777215
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Tool"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":1BD6F
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Update"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":1EC43
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "About"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":21B17
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
      Begin CMC.jcbutton bMenu 
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   181
         Top             =   3000
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         ButtonStyle     =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   " Plugin"
         Mode            =   1
         PictureNormal   =   "frmMain.frx":249EB
         PictureAlign    =   0
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         CaptionAlign    =   0
         TooltipBackColor=   0
         ColorScheme     =   3
      End
   End
   Begin CMC.Downloader Downloader1 
      Left            =   1440
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin CMC.UniDialog UniDialog1 
      Left            =   1680
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   323
      FileCustomFilter=   "frmMain.frx":278BF
      FileDefaultExtension=   "frmMain.frx":278DF
      FileFilter      =   "frmMain.frx":278FF
      FileOpenTitle   =   "frmMain.frx":27947
      FileSaveTitle   =   "frmMain.frx":2797F
      FolderMessage   =   "frmMain.frx":279B7
   End
   Begin VB.Menu mnSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnCScan 
         Caption         =   "Hide Scanner"
      End
      Begin VB.Menu bt0 
         Caption         =   "-"
      End
      Begin VB.Menu mnEPro 
         Caption         =   "Enable Protection"
      End
      Begin VB.Menu mnRun 
         Caption         =   "Run On Startup"
      End
      Begin VB.Menu bt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnUpdate 
         Caption         =   "Update Now"
      End
      Begin VB.Menu btx 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnAct 
      Caption         =   "Action"
      Visible         =   0   'False
      Begin VB.Menu mnFixS 
         Caption         =   "Fix Selected"
      End
      Begin VB.Menu bt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnFixC 
         Caption         =   "Fix Checked"
      End
      Begin VB.Menu mnFixA 
         Caption         =   "Fix All Object"
      End
      Begin VB.Menu bt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnExcL 
         Caption         =   "Add to Exception"
      End
      Begin VB.Menu bt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnExp 
         Caption         =   "Explore Object"
      End
      Begin VB.Menu bt10 
         Caption         =   "-"
      End
      Begin VB.Menu mnProP 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnProses 
      Caption         =   "Proses"
      Visible         =   0   'False
      Begin VB.Menu mnRefresh 
         Caption         =   "Refresh Processes"
      End
      Begin VB.Menu bt5 
         Caption         =   "-"
      End
      Begin VB.Menu mnKillPro 
         Caption         =   "Kill Process"
      End
      Begin VB.Menu mnRestartPro 
         Caption         =   "Restart Process"
      End
      Begin VB.Menu mnPausePro 
         Caption         =   "Pause Process"
      End
      Begin VB.Menu mnResumePro 
         Caption         =   "Resume Process"
      End
      Begin VB.Menu bt6 
         Caption         =   "-"
      End
      Begin VB.Menu mnProProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'CMC * Codenesia Malware Cleaner
'Code   : HrXxX + PmZ
'Desain : Sony
'Tgl : 20 Desember 2009

Dim WithEvents ShellIE As SHDocVw.ShellWindows
Attribute ShellIE.VB_VarHelpID = -1

Dim isCompatch As Boolean
Dim LewatExit  As Boolean
Dim SudahJalan As Boolean ' buffer ajh

Dim IndekPluginTerplih As Long ' buat buffer plugin aj karena gak support unic ListBoxnya
Dim Detik        As Long
Dim Detik2       As Long
Dim Menit        As Long
Dim Jam          As Integer

Dim StatScan      As String ' status scan
Dim UniDialogPath As String ' Penampung Path pada UNIDIALOG
Dim UniDialogFile As String ' Penampung File pada UNIDIALOG

Dim LastPathRClick    As String ' path terakhir yang dikil kanan dari LV2 hasil scan
Dim ListViewClicked   As String ' status nama listview yang di klik kanan

Dim PathDariShellMenu As String
Dim PathLainArr(2)    As String

Dim NewFDMasuk        As String
Dim BufferUpdate      As Long ' penanda sampai mana updatenya

Private Sub bMenu_Click(Index As Integer)
    On Error Resume Next
    Dim i As Byte
    Const t1 = 960, l1 = 2760, h1 = 6495, w1 = 7935
    For i = bMenu.LBound To bMenu.UBound
        If i <> Index Then
            picMenu(i).Visible = False
            bMenu(i).value = False
        End If
    Next
    With picMenu(Index)
        .Top = t1: .Left = l1: .Height = h1: .Width = w1
    End With
    bMenu(Index).value = True
    picMenu(Index).Visible = True
    
    If Index = 3 Then
        txtRetriveInfo.Text = ""
    End If
    
    If Index = 5 Then
       Call EnumPlugin(GetFilePath(App_FullPathW(False)) & "\plugin", lstPlugin)
       IndekPluginTerplih = -1
    End If

End Sub

Private Sub ck12_Click()
If ck12.value = 1 Then
    If Len(j_bahasa(51)) < 1 Then Exit Sub
    Install_CMenu (j_bahasa(51) & " CMC")
Else
   If Len(j_bahasa(51)) < 1 Then Exit Sub
    unInstall_CMenu (j_bahasa(51) & " CMC")
End If
End Sub

Private Sub cmdAddExcFile_Click()
UniDialog1.ShowOpen
If ValidFile(UniDialogFile) = True Then
   ReBuildFileException UniDialogFile, GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile
End If
End Sub

Private Sub cmdAddExcFolder_Click()
UniDialog1.ShowFolder
If Len(UniDialogPath) > 2 Then
   ReBuildPathException UniDialogPath, GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder
End If
End Sub

Private Sub cmdAddVirus_Click()
If txtVirusPath.Text = "" Then
   MsgBox i_bahasa(7), vbExclamation
   Exit Sub
End If
If AddVirusTemp(UniDialogFile, txtVirusName.Text) = True Then
   MsgBox "File [ " & txtVirusPath.Text & " ]" & Chr(13) & _
          i_bahasa(16) & Chr(13) & _
          i_bahasa(17) & " : " & txtVirusName.Text & Chr(13) & Chr(13) & _
          i_bahasa(18) & " !", vbExclamation
   nVirusTmp = nVirusTmp + 1
   frTemp.Caption = c_bahasa(5) & " ( " & nVirusTmp & " ) !"
   lstVirTemp.AddItem txtVirusPath.Text & " - " & txtVirusName.Text
   txtVirusPath.Text = ""
   txtVirusName.Text = ""
   Call bMenu_Click(0)
Else
   MsgBox "File [ " & txtVirusPath.Text & " ]" & Chr(13) & _
          i_bahasa(20) & " !", vbExclamation
   txtVirusPath.Text = ""
   txtVirusName.Text = ""
End If

End Sub

Private Sub cmdApplyLang_Click()
Dim BhasaDipakai As String
    BhasaDipakai = lstLanguage.List(lstLanguage.ListIndex)
    BhasaDipakai = Mid(BhasaDipakai, InStr(BhasaDipakai, "| ") + 2)
    LangUsed = BhasaDipakai
    
    InitLanguange LangUsed
    
    SaveConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    LoadConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    
    MsgBox i_bahasa(15), vbInformation
End Sub

Private Sub cmdBrowse_Click()
    UniDialog1.ShowOpen
    If ValidFile(UniDialogFile) = True Then
       txtVirusPath.Text = UniDialogFile
       txtVirusName.Text = j_bahasa(9) & (nVirusTmp + 1)
    End If
End Sub

Private Sub cmdCancel_Click()
    UniDialogFile = ""
    txtVirusName.Text = ""
    txtVirusPath = ""
End Sub

Private Sub cmdCheckUpdate_Click()
If cmdCheckUpdate.Caption = j_bahasa(27) Then
  bUpdateCompon = False
  HentikanUpdate = False
  BufferUpdate = -1
  
  AmbilUpdateInfo "http://cmc.codenesia.com/files/update/UpdateInfo.txt", GetSpecFolder(USER_DOC) & "\updcmc.tmp"
  cmdCheckUpdate.Caption = j_bahasa(34)
  mnUpdate.Caption = j_bahasa(34)

  lblStatusUpdate.Caption = j_bahasa(33)
Else
  HentikanUpdate = True
  cmdCheckUpdate.Caption = j_bahasa(27)
  mnUpdate.Caption = j_bahasa(27)
  lblStatusUpdate.Caption = j_bahasa(32)
End If
End Sub

Private Sub cmdClearJail_Click()
If MsgBox(i_bahasa(19), vbExclamation + vbYesNo) = vbYes Then
    ClearJail lvJail
End If
End Sub

Private Sub cmdExecutePlug_Click()
If IndekPluginTerplih >= 0 Then
    RunPlugin IndekPluginTerplih
End If
End Sub

Private Sub cmdExplore_Click()
Dim IndekTerpilih   As Long
Dim sPathFile       As String
IndekTerpilih = CariIndekItemTerpilih(lvInfo)
If IndekTerpilih > 0 Then
   sPathFile = lvInfo.ListItems.Item(IndekTerpilih).SubItem(2).Text
   Shell "Explorer.exe /e," & GetFilePath(sPathFile), vbNormalFocus
End If
End Sub

Private Sub cmdFixHidden_Click()
    Call FIX_HIDDEN(lvHidden, BY_CHECKED)
End Sub

Private Sub cmdFixHiddenAll_Click()
    Call FIX_HIDDEN(lvHidden, BY_ALL)
End Sub

Private Sub cmdFixMalware_Click()
    cmdFixMalware.Enabled = False
    cmdFixMalwareAll.Enabled = False
    FiX_Malware lvMalware, BY_CHECKED, 16
    cmdFixMalwareAll.Enabled = True
    cmdFixMalware.Enabled = True
End Sub

Private Sub cmdFixMalwareAll_Click()
    cmdFixMalware.Enabled = False
    cmdFixMalwareAll.Enabled = False
    Call FiX_Malware(lvMalware, BY_ALL, 16)
    cmdFixMalwareAll.Enabled = True
    cmdFixMalware.Enabled = True
End Sub

Private Sub cmdFixReg_Click()
    cmdFixRegAll.Enabled = False
    cmdFixReg.Enabled = False
    Call FiX_REGISTRY(lvRegistry, BY_CHECKED)
    cmdFixRegAll.Enabled = True
    cmdFixReg.Enabled = True
End Sub

Private Sub cmdFixRegAll_Click()
    cmdFixRegAll.Enabled = False
    cmdFixReg.Enabled = False
    Call FiX_REGISTRY(lvRegistry, BY_ALL)
    cmdFixRegAll.Enabled = True
    cmdFixReg.Enabled = True
End Sub

Private Sub cmdKillPris_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long
Dim Counter         As Long

If MsgBox(i_bahasa(21) & " ?", vbExclamation + vbYesNo) = vbYes Then
   For Counter = 1 To lvJail.ListItems.Count
       If lvJail.ListItems.Item(Counter).Selected = False Then GoTo LBL_LANJUT
       PrisName = lvJail.ListItems.Item(Counter).SubItem(4).Text
       KillPrisonner PrisName, lvJail
LBL_LANJUT:
   Next
Call READ_DATA_JAIL(FolderJail)
End If
End Sub

Private Sub cmdMoreInfo_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.cmc.codenesia.com", vbNullString, "C:\", 1
End Sub

Private Sub cmdProperties_Click()
Dim IndekTerpilih   As Long
IndekTerpilih = CariIndekItemTerpilih(lvInfo)
If IndekTerpilih > 0 Then
   ShowProperties lvInfo.ListItems.Item(IndekTerpilih).SubItem(2).Text, Me.hwnd
End If
End Sub

Private Sub cmdRelease_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long

IndekTerpilih = CariIndekItemTerpilih(lvJail)
If IndekTerpilih > 0 Then
   If MsgBox(i_bahasa(22) & " ?", vbExclamation + vbYesNo) = vbYes Then
      ReleasePrisoner lvJail.ListItems.Item(IndekTerpilih).SubItem(2).Text, lvJail.ListItems.Item(IndekTerpilih).SubItem(4).Text, lvJail
   End If
End If
End Sub

Private Sub cmdReleaseTo_Click()
Dim PrisName        As String
Dim IndekTerpilih   As Long
Dim PrisonerFName  As String

IndekTerpilih = CariIndekItemTerpilih(lvJail)
UniDialog1.ShowFolder
    If PathIsDirectory(StrPtr(UniDialogPath)) <> 0 Then
       If IndekTerpilih > 0 Then
          PrisonerFName = GetFileName(lvJail.ListItems.Item(IndekTerpilih).SubItem(2).Text)
          If MsgBox(i_bahasa(22) & " here ?", vbExclamation + vbYesNo) = vbYes Then
             ReleasePrisoner UniDialogPath & "\" & PrisonerFName, lvJail.ListItems.Item(IndekTerpilih).SubItem(4).Text, lvJail
          End If
       End If
    End If
End Sub

Private Sub cmdRemExcFile_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\File.lst"
    ReadExceptFile GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile
End Sub

Private Sub cmdRemExcFile1_Click()
    RemoveExceptionByIndek lstExceptFile.ListIndex, FILE_EXC
    JumFileExcep = ReadExceptFile(GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile)
End Sub

Private Sub cmdRemExcReg_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\Reg.lst"
    ReadExceptReg GetFilePath(App_FullPathW(False)) & "\Reg.lst", lstExceptReg
End Sub

Private Sub cmdRemExcReg1_Click()
    RemoveExceptionByIndek lstExceptReg.ListIndex, REG_EXC
    JumRegExcep = ReadExceptReg(GetFilePath(App_FullPathW(False)) & "\Reg.lst", lstExceptReg)
End Sub

Private Sub cmdRemovePath_Click()
    HapusFile GetFilePath(App_FullPathW(False)) & "\Path.lst"
    ReadExceptPath GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder
End Sub

Private Sub cmdRemovePath1_Click()
RemoveExceptionByIndek lstExceptFolder.ListIndex, PATH_EXC
JumPathExcep = ReadExceptPath(GetFilePath(App_FullPathW(False)) & "\Path.lst", lstExceptFolder)
End Sub

Private Sub cmdSave_Click()
    SaveConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    LoadConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    MsgBox i_bahasa(15), vbInformation
End Sub

Private Sub cmdStartScan_Click()
Dim lstCek      As Collection
Dim iCount      As Long

'init
StatScan = d_bahasa(17)
Set lstCek = New Collection
DirTree.OutPutPath lstCek

If cmdStartScan.Caption = a_bahasa(5) Then

   If MaulanjutScan(lvMalware) = False Then Exit Sub
   
   Call ResetObjek
   
   If Len(NewFDMasuk) = 3 Then GoTo LBL_SCAN_FD Else GoTo END_LBL_FD

LBL_SCAN_FD: ' scan FD masuk
  'mulai buffer path yang akan di scan
  BufferPath NewFDMasuk, True
  cmdStartScan.Caption = a_bahasa(7)
  ' init Progress Bar
  PB1.value = 0
  PB1.Max = FileToScan
  
  lbStatus.Caption = d_bahasa(15)
  KumpulkanFile NewFDMasuk, lbObject, True, True
  GoTo LBL_LOMPATAN_FD
END_LBL_FD:
   
   If WinNode = True Then PathLainArr(0) = GetSpecFolder(WINDOWS_DIR) Else PathLainArr(0) = ""
   If DocNode = True Then PathLainArr(1) = GetSpecFolder(USER_DOC) Else PathLainArr(1) = ""
   If ProgNode = True Then PathLainArr(2) = GetSpecFolder(PROGRAM_FILE) Else PathLainArr(2) = ""
   
   cmdStartScan.Caption = a_bahasa(6)
   
   'mulai buffer path yang akan di scan
   For iCount = 1 To lstCek.Count
      If WithBuffer = False Then Exit For
      BufferPath lstCek(iCount), True
   Next
   iCount = 0
   
   'buffer path tambahan
   For iCount = 0 To 2
      If WithBuffer = False Then Exit For
      If Len(PathLainArr(iCount)) > 0 Then BufferPath PathLainArr(iCount), True
   Next
   'buffer buat shell menu (jika folder)
   If Len(PathDariShellMenu) > 0 And ValidFile(PathDariShellMenu) = False Then BufferPath PathDariShellMenu, True

   
   cmdStartScan.Caption = a_bahasa(7)
   
    
   If RegNode = True Then
      lbStatus.Caption = d_bahasa(10)
      If ck3.value = 1 Then 'dengan auto detect useles value
         ScanRegistry lbObject, False, True
      Else
         ScanRegistry lbObject, False, False
      End If
   End If
   
   If ProsesNode = True Then 'scan proses + service
      lbStatus.Caption = d_bahasa(11)
      Call ScanService(lbObject, True)
      
      lbStatus.Caption = d_bahasa(12)
      Call ScanProses(False, lbObject)
   
      lbStatus.Caption = j_bahasa(0) ' module
      Call ScanProses(True, lbObject)

   End If
   
   If StartUpNode = True Then ' scan startup
      lbStatus.Caption = d_bahasa(13)
      ScanRegStartup lbObject, True
   End If
   
   If UCase$(Left$(Command, 2)) <> "-S" Then lbStatus.Caption = d_bahasa(14) ' root drive
   If UCase$(Left$(Command, 2)) <> "-S" Then ScanRootDrive lbObject

   ' init Progress Bar
   PB1.value = 0
   PB1.Max = FileToScan
   
   ' reset
   iCount = 0
   ' Mulai pindai
   lbStatus.Caption = d_bahasa(15)
     
   'path lain dulu
   For iCount = 0 To 2
       If BERHENTI = True Then Exit For
       If Len(PathLainArr(iCount)) > 0 Then KumpulkanFile PathLainArr(iCount), lbObject, True, True
   Next
   'scan dari path shellmenu
   If Len(PathDariShellMenu) > 0 Then
      If ValidFile(PathDariShellMenu) = False Then ' jika folder
         KumpulkanFile PathDariShellMenu, lbObject, True, True
      Else ' jika file
         CocokanDataBase PathDariShellMenu
         FileCheck = FileCheck + 1
         FileFound = FileFound + 1
         lbFileCheck.Caption = ": " & Right$("00000000" & FileCheck, 8)
         lbFileFound.Caption = ": " & Right$("00000000" & FileFound, 8)
      End If
   End If
   ' reset
   iCount = 1
   For iCount = 1 To lstCek.Count
       If BERHENTI = True Then Exit For
       KumpulkanFile lstCek(iCount), lbObject, True, True
   Next

LBL_LOMPATAN_FD:
   tmTime.Enabled = False
   cmdStartScan.Caption = a_bahasa(5)
   If StatScan = d_bahasa(16) And WithBuffer = True Then StatScan = d_bahasa(16) & " !"
   lbStatus.Caption = StatScan
   
   BERHENTI = True
   
   MsgBoxU String(6, ChrW$(&H20AA)) & " " & j_bahasa(41) & " " & String(6, ChrW$(&H20AA)) & Chr(13) & _
      "_________________________________" & Chr(13) & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(42) & ": " & StatScan & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(43) & ": " & FileFound & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(44) & ": " & FileCheck & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(45) & ": " & FileNotCheck & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(46) & ": " & VirusFound & Chr(13) & ChrW$(&H221A) & _
      " " & d_bahasa(8) & ": " & InfoFound & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(47) & ": " & nRegVal & " " & j_bahasa(12) & Chr(13) & ChrW$(&H221A) & _
      " " & j_bahasa(48) & ": " & nErrorReg & " " & j_bahasa(12) & Chr(13) & _
      "_________________________________" & Chr(13) & Chr(13) & _
      j_bahasa(49) & " :: " & Now, ChrW$(&H20AA) & " C.M.C " & ChrW$(&H20AA), 0, Me
      
      TabMain.GantiJudul 2, b_bahasa(1) & "-" & VirusFound
      TabMain.GantiJudul 3, b_bahasa(2) & "-" & nErrorReg
      TabMain.GantiJudul 4, b_bahasa(3) & "-" & lvHidden.ListItems.Count
      TabMain.GantiJudul 5, b_bahasa(4) & "-" & lvInfo.ListItems.Count
      
      Call ReBack ' Aktifkan yang peru diaktifkan
      
      Me.Refresh
      
      Me.WindowState = vbNormal
      Me.Show
      If lvMalware.ListItems.Count > 0 Then TabMain.AktifTab = 2
      
      PathDariShellMenu = "" 'kosongkan lagi yang dari shell menu
      NewFDMasuk = "" ' FD masuk kosong
      
      If ck10.value = 1 Then tmFlash.Enabled = True

ElseIf cmdStartScan.Caption = a_bahasa(6) Then
   Call StopKumpulkan
   WithBuffer = False
Else
   Call StopKumpulkan
   BERHENTI = True
   StatScan = d_bahasa(16)
   cmdStartScan.Caption = a_bahasa(5)
   PathDariShellMenu = "" 'kosongkan lagi yang dari shell menu
End If

End Sub


Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String)
   BufferUpdate = BufferUpdate + 1
   tmUpdate.Enabled = True
End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)
    HentikanUpdate = True
    lblStatusUpdate.Caption = "Error Download..."
End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    PBC.Max = MaxBytes
    PBC.value = CurBytes
End Sub

Private Sub Form_DblClick()
'frmCeksumIcon.Show
End Sub

Private Sub Form_Load()
If App.PrevInstance = True And UCase$(Left$(Command, 2)) <> "-S" Then
   MsgBox "Proses yang sama masih berjalan !", vbExclamation
   End
End If

Call InitAplikasi

Call bMenu_Click(0)

If UCase$(Left$(Command, 2)) <> "-S" Then
   BERHENTI = False
   ScanProses False, lbObject
   If VirusFound > 0 Then
      MsgBox "CMC " & j_bahasa(59) & " !", vbOKOnly + vbExclamation
   End If
   
   Call ReBack
   BERHENTI = True
   
End If
lbObject.Caption = ""

If UCase$(Left$(Command, 2)) <> "-S" Then
  Call CompactObject
  Call LayOnDekstop
End If

tmAwal.Enabled = True

SudahJalan = True ' informasikan sopware sudah jalan
IndekPluginTerplih = -1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lHasil  As Long
Dim HorX    As Long
    
    If Me.ScaleMode = vbPixels Then
        HorX = x
    Else
        HorX = x / Screen.TwipsPerPixelX
    End If
    
    Select Case HorX
        Case WM_LBUTTONDBLCLK
            Me.WindowState = vbNormal
            lHasil = SetForegroundWindow(Me.hwnd)
            mnCScan.Caption = g_bahasa(0)
            Me.Show
        Case WM_RBUTTONUP 'Tampilkan menu Popup saat klik kanan.
            lHasil = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mnSystray
    End Select

End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then
       mnCScan.Caption = g_bahasa(1)
    End If
    Me.Height = 7950
    Me.Width = 10830
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LewatExit = False And Left$(Command, 2) <> "-S" Then ' klo dari shell scan tanda X juga berguna nutup
    Cancel = 1
    Shell_NotifyIcon NIM_DELETE, nID
    Call UpdateIcon(Me.Icon, "CMC PH#3.5 - CodenesiaSoft", Me)
    Me.WindowState = vbMinimized
    Me.Hide
ElseIf BERHENTI = False Then
    Cancel = 1
    MsgBox i_bahasa(8), vbExclamation
Else
    Me.Show
    Shell_NotifyIcon NIM_DELETE, nID
    Unload frmRTP
    End
End If
End Sub

Private Sub lblWeb_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.codenesia.com", vbNullString, "C:\", 1
End Sub

Private Sub lblWeb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblWeb.ForeColor = vbYellow
    lblWeb.FontUnderline = True
End Sub


Private Sub lstExceptFile_Click()
    lstExceptFile.ToolTipText = lstExceptFile.List(lstExceptFile.ListIndex)
End Sub

Private Sub lstExceptFolder_Click()
    lstExceptFolder.ToolTipText = lstExceptFolder.List(lstExceptFolder.ListIndex)
End Sub

Private Sub lstExceptReg_Click()
    lstExceptReg.ToolTipText = lstExceptReg.List(lstExceptReg.ListIndex)
End Sub

Private Sub lstLanguage_Click()
WriteLngInfoToLabel lstLanguage.List(lstLanguage.ListIndex), lblLangID1, lblLangSel1, lblLangAut1

End Sub

Private Sub lstModule_DblClick()
If Left$(lstModule.List(lstModule.ListIndex), 2) = "0x" Then
    If MsgBox(i_bahasa(9), vbExclamation + vbYesNo) = vbYes Then
        Call UnloadModuleForce(lstModule.List(lstModule.ListIndex), lstModule, lblSelectedPID.Caption)
    End If
End If
End Sub

Private Sub lstPlugin_Click()
Dim MyIndek As Long

MyIndek = lstPlugin.ListIndex
If MyIndek >= 0 Then
   IndekPluginTerplih = MyIndek
   RetrievePlugInfo MyIndek, lstPlugin, lblPlugSelect1, lblPlugAut1, lblPlugAutEmail1, lblPlugAutSite1, lblPlugVer1, lblPlugDesc1
End If
End Sub

Private Sub lstVirTemp_Click()
    lstVirTemp.ToolTipText = lstVirTemp.List(lstVirTemp.ListIndex)
End Sub

Private Sub lvHidden_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvHidden_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
  If iButton = vbccMouseRButton Then
      LastPathRClick = oItem.SubItem(2).Text ' berikan info klik kanan terkahir
      ListViewClicked = "HIDDEN"
      mnExcL.Enabled = False
      If ValidFile(LastPathRClick) = True Then
         mnProP.Enabled = True
         mnExp.Enabled = True
      Else
         mnProP.Enabled = False
         mnExp.Enabled = False
      End If
      PopupMenu mnAct, 0, , , mnFixS
  Else
      'If oItem.Checked = True Then oItem.Checked = False Else oItem.Checked = True
      picInfoHidden.Cls
      RetrieveIcon oItem.SubItem(2).Text, picInfoHidden, ricnLarge
  End If
End Sub

Private Sub lvInfo_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvInfo_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
  If iButton = vbccMouseRButton Then
     LastPathRClick = oItem.SubItem(2).Text ' berikan info klik kanan terkahir
  Else
     picIconInfo.Cls
     RetrieveIcon oItem.SubItem(2).Text, picIconInfo, ricnLarge
  End If
End Sub

Private Sub lvJail_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvMalware_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvMalware_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
  If iButton = vbccMouseRButton Then
     LastPathRClick = oItem.SubItem(2).Text ' berikan info klik kanan terkahir
     ListViewClicked = "MALWARE"
     If ValidFile(LastPathRClick) = True Then
         mnProP.Enabled = True
         mnExp.Enabled = True
         mnExcL.Enabled = True
     Else
         mnProP.Enabled = False
         mnExp.Enabled = False
         mnExcL.Enabled = False
     End If

     PopupMenu mnAct, 0, , , mnFixS
   Else
      'If oItem.Checked = True Then oItem.Checked = False Else oItem.Checked = True
      picInfoMalware.Cls
      RetrieveIcon oItem.SubItem(2).Text, picInfoMalware, ricnLarge
   End If
End Sub

Private Sub lvProses_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvProses_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
    Call ENUM_MODULE(oItem.SubItem(3).Text, lstModule)
    lblSelectedPID.Caption = oItem.SubItem(3).Text
If iButton = vbccMouseRButton Then PopupMenu mnProses, 0, , , mnRefresh
End Sub

Private Sub lvRegistry_ColumnClick(ByVal oColumn As cColumn)
    oColumn.Sort
End Sub

Private Sub lvRegistry_ItemClick(ByVal oItem As cListItem, ByVal iButton As evbComCtlMouseButton)
  If iButton = vbccMouseRButton Then
     LastPathRClick = oItem.SubItem(2).Text ' berikan info klik kanan terkahir
     ListViewClicked = "REGISTRY"
     
     mnProP.Enabled = False
     mnExp.Enabled = False
     mnExcL.Enabled = True
     
     PopupMenu mnAct, 0, , , mnFixS
   Else
      'If oItem.Checked = True Then oItem.Checked = False Else oItem.Checked = True
   End If
End Sub

Private Sub mnCScan_Click()
If mnCScan.Caption = g_bahasa(1) Then
    Me.WindowState = vbNormal
    Me.Show
    mnCScan.Caption = g_bahasa(0)
Else
    mnCScan.Caption = g_bahasa(1)
    Me.WindowState = vbMinimized
    Me.Hide
End If
End Sub

Private Sub mnEPro_Click()
    If mnEPro.Checked = True Then
       mnEPro.Checked = False
       ck8.value = 0
    Else
       mnEPro.Checked = True
       ck8.value = 1
    End If
    SaveConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    LoadConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    
    If StatusRTP = True Then
       TampilkanBalon frmMain, i_bahasa(24) & " !", i_bahasa(26), NIIF_INFO
    Else
       TampilkanBalon frmMain, i_bahasa(25) & " !", i_bahasa(27), NIIF_WARNING
    End If
    
    Sleep 2000
    
    CabutBalon Me

End Sub

Private Sub mnExcL_Click()
Dim nIndek As Long
nIndek = CariIndekItemTerpilih(lvMalware)
If ValidFile(LastPathRClick) = True Then
   ReBuildFileException LastPathRClick, GetFilePath(App_FullPathW(False)) & "\File.lst", lstExceptFile
   If nIndek > 0 Then lvMalware.ListItems.Remove nIndek
Else ' berarti exception untuk registry
   nIndek = CariIndekItemTerpilih(lvRegistry)
   LastPathRClick = GetKeyPathAndValueClean(LastPathRClick) ' dihilangkan klo ada tandaa seperti "=>"
   ReBuildRegException LastPathRClick, GetFilePath(App_FullPathW(False)) & "\Reg.lst", lstExceptReg
   If nIndek > 0 Then lvRegistry.ListItems.Remove nIndek

End If
End Sub

Private Sub mnExit_Click()
    Call LepasSemuaKunci
    LewatExit = True
    Set ShellIE = Nothing
    SaveConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    'Unload frmRTP
    Unload Me
End Sub

Private Sub mnExp_Click()
   Shell "Explorer.exe /e," & GetFilePath(LastPathRClick), vbNormalFocus
End Sub

Private Sub mnFixA_Click()
Select Case ListViewClicked
    Case "MALWARE":  FiX_Malware lvMalware, BY_ALL, 16
    Case "REGISTRY": FiX_REGISTRY lvRegistry, BY_ALL
    Case "HIDDEN":   FIX_HIDDEN lvHidden, BY_ALL
End Select
End Sub

Private Sub mnFixC_Click()
Select Case ListViewClicked
    Case "MALWARE":  FiX_Malware lvMalware, BY_CHECKED, 16
    Case "REGISTRY": FiX_REGISTRY lvRegistry, BY_CHECKED
    Case "HIDDEN":   FIX_HIDDEN lvHidden, BY_CHECKED
End Select
End Sub

Private Sub mnFixS_Click()
Select Case ListViewClicked
    Case "MALWARE":  FiX_Malware lvMalware, BY_SELECT, 16
    Case "REGISTRY": FiX_REGISTRY lvRegistry, BY_SELECT
    Case "HIDDEN":   FIX_HIDDEN lvHidden, BY_SELECT
End Select
End Sub

Private Sub mnKillPro_Click()
Dim nIndek As Long
Dim PID    As Long
Dim spath  As String

nIndek = CariIndekItemTerpilih(lvProses)

If nIndek > 0 Then
   PID = CLng(lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   spath = lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If KillProses(PID, spath, False, True) = True Then
      MsgBox i_bahasa(10) & ": [" & PID & "] " & i_bahasa(11)
      Call ENUM_PROSES(lvProses, picBuffer)
   Else
      MsgBox i_bahasa(10) & ": [" & PID & "] " & i_bahasa(12)
   End If
End If

End Sub

Private Sub mnPausePro_Click()
Dim nIndek As Long
Dim PID    As Long

nIndek = CariIndekItemTerpilih(lvProses)
If nIndek > 0 Then
   PID = CLng(lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   lvProses.ListItems.Item(nIndek).SubItem(5).Text = SuspendProses(PID, True)
End If
End Sub

Private Sub mnProP_Click()
If ValidFile(LastPathRClick) = True Then
   ShowProperties LastPathRClick, Me.hwnd
End If
End Sub

Private Sub mnProProperties_Click()
Dim spath  As String
Dim nIndek As Long

nIndek = CariIndekItemTerpilih(lvProses)
If nIndek > 0 Then
   spath = lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If ValidFile(spath) = True Then ShowProperties spath, Me.hwnd
End If
End Sub

Private Sub mnRefresh_Click()
    Call ENUM_PROSES(lvProses, picBuffer) ' Refresh
End Sub

Private Sub mnRestartPro_Click()
Dim nIndek As Long
Dim PID    As Long
Dim spath  As String

nIndek = CariIndekItemTerpilih(lvProses)

If nIndek > 0 Then
   PID = CLng(lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   spath = lvProses.ListItems.Item(nIndek).SubItem(9).Text
   If KillProses(PID, spath, True, False) = True Then
      MsgBox i_bahasa(10) & ": [" & PID & "] " & i_bahasa(13)
      Call ENUM_PROSES(lvProses, picBuffer)
   Else
      MsgBox i_bahasa(10) & ": [" & PID & "] " & i_bahasa(14)
   End If
End If

End Sub

Private Sub mnResumePro_Click()
Dim nIndek As Long
Dim PID    As Long

nIndek = CariIndekItemTerpilih(lvProses)
If nIndek > 0 Then
   PID = CLng(lvProses.ListItems.Item(nIndek).SubItem(3).Text)
   lvProses.ListItems.Item(nIndek).SubItem(5).Text = SuspendProses(PID, False)
End If
End Sub

Private Sub mnRun_Click()
    If mnRun.Checked = True Then
       mnRun.Checked = False
       ck7.value = 0
    Else
       mnRun.Checked = True
       ck7.value = 1
    End If
    SaveConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
    LoadConfig GetFilePath(App_FullPathW(False)) & "\CMC.ini"
End Sub


Private Sub mnUpdate_Click()
    Me.WindowState = vbNormal
    Me.Show
    Call bMenu_Click(3)
    Call cmdCheckUpdate_Click
End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index = 0 Then ' pic header
   lblWeb.ForeColor = vbRed
   lblWeb.FontUnderline = False
End If
End Sub

Private Sub rtp_mode1_PathChange(Index As Integer, strPath As String)
If StatusRTP = True And BERHENTI = True Then
    ScanPatWithRTP strPath
End If
End Sub

Private Sub TabAbout_Click(Index As Integer)
Select Case Index
       'Case 1: Picture1(1).Left = 120
       Case 2: If SudahJalan = True Then Call ListVirus(lstListWorm)  ': PicCmcInfo.Left = 120
End Select
End Sub

Private Sub TabTool_Click(Index As Integer)
Select Case Index
       Case 1: If SudahJalan = True Then Call ENUM_PROSES(lvProses, picBuffer) ': PicProsesManager.Left = 120 ' Proses Manager
       'Case 2: picTempMalware.Left = 120 ' Tab Malware Signer
       Case 3: If SudahJalan = True Then Call READ_DATA_JAIL(FolderJail) ': PicJailCont.Left = 120
End Select

End Sub

Private Sub tmAwal_Timer()
Dim nFS As Long
Detik2 = Detik2 + 1
If Detik2 = 1 Then
Select Case Left$(Command, 2)
  Case "-A" ' dari auto run
      Me.WindowState = vbMinimized
      Me.Hide
      Call InitAplikasi2
      If StatusRTP = True Then
         TampilkanBalon frmMain, i_bahasa(24) & " !", i_bahasa(26), NIIF_INFO
      Else
         TampilkanBalon frmMain, i_bahasa(25) & " !", i_bahasa(27), NIIF_WARNING
      End If
  Case "-S" ' scan dari context menu
      Me.WindowState = vbNormal
      PathDariShellMenu = Mid$(Command, 4)
      nFS = EnumFileSystem
      If nFS = 0 Then
         TampilkanBalon frmMain, i_bahasa(23) & " !", i_bahasa(27), NIIF_WARNING
      End If
      Call cmdStartScan_Click
      tmAwal.Enabled = False
  Case Else ' double klik
      Me.WindowState = vbNormal
      Call InitAplikasi2
End Select
End If

If Detik2 = 8 Then
   CabutBalon Me
End If

If Detik2 = 12 Then ' klo mau cek online update
   If ck9.value = 1 Then Call AmbilUpdateInfo("http://cmc.codenesia.com/files/update/UpdateInfo.txt", GetSpecFolder(USER_DOC) & "\updcmc.tmp")
   HentikanUpdate = True
End If

If Detik2 = 22 Then ' bayangkan aja udah selesai ambil informasinya
   If ck9.value = 1 Then Call CheckUpdate(GetSpecFolder(USER_DOC) & "\updcmc.tmp", lblStatusUpdate)
   tmAwal.Enabled = False
   If bUpdateCompon = True Then
      CabutBalon Me
      If MsgBox("CMC menemukan file update baru pada server !" & Chr(13) & _
                "Akankah anda mengupdate database CMC anda sekarang ?", vbYesNo + vbInformation, "Update") = vbYes Then
                Me.WindowState = vbNormal
                Me.Show
                Call bMenu_Click(3)
                Call cmdCheckUpdate_Click
      End If
   End If
End If

End Sub

Private Sub tmFlash_Timer()
Dim sDriveName          As String
Dim DriveLabel          As String
Dim nDriveNameLen       As Long

If BERHENTI = False Or UCase$(Left$(Command, 2)) = "-S" Then ' jika scan scan jalan
   tmFlash.Enabled = False
   Exit Sub
End If
If AdakahFDBaru(LastFlashVolume) = True Then
   nDriveNameLen = 128
   sDriveName = String$(nDriveNameLen, 0)
   If GetVolumeInformationW(StrPtr(Chr(LastFlashVolume) & ":\"), StrPtr(sDriveName), nDriveNameLen, ByVal 0, ByVal 0, ByVal 0, 0, 0) Then
       DriveLabel = Left$(sDriveName, InStr(1, sDriveName, ChrW$(0)) - 1)
   Else
       DriveLabel = vbNullString
   End If
   Call BuilDirTree ' refresh dir tree
   TampilkanBalon Me, j_bahasa(60) & " [ " & DriveLabel & " (" & Chr(LastFlashVolume) & ") ] !", i_bahasa(26), NIIF_INFO
   If MsgBox("CMC " & j_bahasa(61) & " CMC" & " [ " & DriveLabel & " (" & Chr(LastFlashVolume) & ") ] ?", vbYesNo + vbExclamation) = vbYes Then
      NewFDMasuk = Chr(LastFlashVolume) & ":\"
      TampilkanBalon Me, j_bahasa(62) & " [ " & DriveLabel & " (" & Chr(LastFlashVolume) & ") ] !", i_bahasa(26), NIIF_INFO
      Call cmdStartScan_Click
   End If
End If
End Sub

Private Sub tmTime_Timer()
Detik = Detik + 1
If Detik = 60 Then
   Detik = 0
   Menit = Menit + 1
End If
If Menit = 60 Then
   Menit = 0
   Jam = Jam + 1
End If
lbTime.Caption = ": " & Jam & " :" & Menit & " :" & Detik
End Sub


Private Sub tmUpdate_Timer()
Dim TmpPath  As String
Dim TmpPath2 As String
Dim MyPath   As String

If HentikanUpdate = True Then GoTo LBL_MATI_AJ

TmpPath = GetSpecFolder(USER_DOC) & "\updcmc.tmp"
TmpPath2 = GetSpecFolder(USER_DOC) & "\cmctmp.txt"
MyPath = GetFilePath(App_FullPathW(False))

Select Case BufferUpdate
    Case 0 ' baru ambil updateinfo.txt
         txtRetriveInfo.Text = CheckUpdate(TmpPath, lblStatusUpdate)
         If bUpdateCompon = True Then ' berarti ada update terbaru
            mnUpdate.Caption = j_bahasa(34)
            cmdCheckUpdate.Caption = j_bahasa(34)
            UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
         End If
    Case 1 ' selesai update komponen db-0 (0x.cmc)
         MoveIfValidComp TmpPath2, MyPath & "\sign\" & Hex$(BufferUpdate - 1) & "x.cmc", lblStatusUpdate
         UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
    Case Is < 16 ' selesai update komponen db-1 (1x.cmc) dst
         MoveIfValidComp TmpPath2, MyPath & "\sign\" & Hex$(BufferUpdate - 1) & "x.cmc", lblStatusUpdate
         UpdateKomponen PB_UPD, lblStatusUpdate, BufferUpdate
    Case 16 ' selesai update komponen db-15 (terakhir PE)
         MoveIfValidComp TmpPath2, MyPath & "\sign\" & Hex$(BufferUpdate - 1) & "x.cmc", lblStatusUpdate
         UpdateKomponenNonPE PB_UPD, lblStatusUpdate, BufferUpdate - 16
    Case Is < 32
         MoveIfValidComp TmpPath2, MyPath & "\signx\" & Hex$(BufferUpdate - 17) & "z.cmc", lblStatusUpdate
         UpdateKomponenNonPE PB_UPD, lblStatusUpdate, BufferUpdate - 16
    Case 32 ' selsai sampai akhir
         MoveIfValidComp TmpPath2, MyPath & "\signx\" & Hex$(BufferUpdate - 17) & "z.cmc", lblStatusUpdate
         cmdCheckUpdate.Caption = j_bahasa(27)
         mnUpdate.Caption = j_bahasa(27)
         lblStatusUpdate.Caption = j_bahasa(31)
         PBC.value = PBC.Max
         tmUpdate.Enabled = False
         BufferUpdate = -1
         Call BacaDatabase
         Call ListVirus(lstListWorm)
         Exit Sub
End Select
PBC.value = 0
tmUpdate.Enabled = False ' matiin lagi...
Exit Sub

LBL_MATI_AJ:
   PBC.value = 0
   tmUpdate.Enabled = False
   cmdCheckUpdate.Caption = j_bahasa(27)
   mnUpdate.Caption = j_bahasa(27)
End Sub

Private Sub UniDialog1_FolderCancel(ByVal CancelType As UniDialogFolderCancel)
    UniDialogPath = ""
End Sub

Private Sub UniDialog1_FolderSelect(ByVal Path As String)
    UniDialogPath = Path
End Sub

Private Sub UniDialog1_OpenCancel(ByVal CancelType As UniDialogFileCancel)
    UniDialogFile = ""
End Sub

Private Sub UniDialog1_OpenFile(ByVal Filename As String)
    UniDialogFile = Filename
End Sub


Private Sub ResetObjek()
   lbMalware.Caption = ": 000000 " & d_bahasa(38)
   lbReg.Caption = ": 000000 " & d_bahasa(38)
   lbHidden.Caption = ": 000000 " & d_bahasa(38)
   lbInfo.Caption = ": 000000 " & d_bahasa(38)
   lbBypass.Caption = ": 00000000"
   lbFileFound.Caption = ": 00000000"
   lbFileCheck.Caption = ": 00000000"
   
   lvMalware.ListItems.Clear
   lvRegistry.ListItems.Clear
   lvHidden.ListItems.Clear
   lvInfo.ListItems.Clear
   
   TabMain.GantiJudul 2, b_bahasa(1)
   TabMain.GantiJudul 3, b_bahasa(2)
   TabMain.GantiJudul 4, b_bahasa(3)
   TabMain.GantiJudul 5, b_bahasa(4)

   
   PB1.value = 0
   VirusFound = 0
   FileFound = 0
   FileCheck = 0
   FileNotCheck = 0
   nRegVal = 0
   nErrorReg = 0
   FileToScan = 0
   InfoFound = 0
   
   Detik = 0
   Menit = 0
   Jam = 0
   
   WithBuffer = True ' nilai awal true
   BERHENTI = False
   
   Picture2(0).Enabled = False
   Picture3.Enabled = False
   Picture4.Enabled = False
   Picture5.Enabled = False
   
   cmdFixMalware.Enabled = False
   cmdFixMalwareAll.Enabled = False
   cmdFixReg.Enabled = False
   cmdFixRegAll.Enabled = False
   cmdFixHidden.Enabled = False
   cmdFixHiddenAll.Enabled = False
   cmdExplore.Enabled = False
   cmdProperties.Enabled = False
   
   tmTime.Enabled = True
   
   Call LepasSemuaKunci
   
   DataAutorun = "" ' Reset
   TargetShorcutOnFD = "" ' Reset
End Sub

Private Sub ReBack()
   Picture2(0).Enabled = True
   Picture3.Enabled = True
   Picture4.Enabled = True
   Picture5.Enabled = True
   
   If lvMalware.ListItems.Count > 0 Then
      cmdFixMalware.Enabled = True
      cmdFixMalwareAll.Enabled = True
   End If
   If lvRegistry.ListItems.Count > 0 Then
      cmdFixReg.Enabled = True
      cmdFixRegAll.Enabled = True
   End If
   If lvHidden.ListItems.Count > 0 Then
      cmdFixHidden.Enabled = True
      cmdFixHiddenAll.Enabled = True
   End If
   If lvInfo.ListItems.Count > 0 Then
      cmdExplore.Enabled = True
      cmdProperties.Enabled = True
   End If
   
End Sub

'....................................... Sudah Mulai RTP

Private Sub ShellIE_WindowRegistered(ByVal lCookie As Long) ' user membuka explorer baru
    If StatusRTP = True Then Call MulaiRTP ' jika TRUE ajh
End Sub

' Mulai RTP
Private Sub MulaiRTP()
On Error Resume Next
If isCompatch = False Then
   Dim i As Integer, CNT As Integer
   CNT = ShellIE.Count - 1
   For i = 0 To CNT
       If (rtp_mode1.Count - 1) < CNT Then
          AddIEObj i
       End If
          If FindID(ShellIE(i).hwnd) = False Then
             rtp_mode1(i).EnabledMonitoring True
             rtp_mode1(i).AddSubClass ShellIE(i)
          End If
   Next i
End If
End Sub

Sub AddIEObj(Index As Integer)
On Error GoTo salah
    Load rtp_mode1(Index)
salah:
End Sub

Function FindID(id As Long) As Boolean
On Error GoTo salah
    Dim i As Integer
    For i = 0 To rtp_mode1.Count - 1
        If rtp_mode1(i).IEKey = id Then
           FindID = True
        End If
    Next i
salah:
End Function

Private Sub CompactObject() ' untuk aktifkan rtp
On Error Resume Next
isCompatch = True
   Dim i As Integer, CNT As Integer
   For i = 0 To rtp_mode1.Count - 1
       rtp_mode1(i).SetIENothing
   Next i
       
   Set ShellIE = Nothing
   For i = 1 To rtp_mode1.Count - 1
        Unload rtp_mode1(i)
   Next i
   
   Set ShellIE = New SHDocVw.ShellWindows
   CNT = ShellIE.Count - 1
   For i = 0 To CNT
       If i > 0 Then
          AddIEObj i
       End If
          rtp_mode1(i).AddSubClass ShellIE(i)
   Next i
isCompatch = False
End Sub


