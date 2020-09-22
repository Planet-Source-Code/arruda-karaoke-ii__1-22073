VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00456B81&
   BorderStyle     =   0  'None
   Caption         =   "SoftKar"
   ClientHeight    =   8685
   ClientLeft      =   270
   ClientTop       =   150
   ClientWidth     =   11400
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   Begin VB.CheckBox Check4 
      BackColor       =   &H00446C82&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   4995
      TabIndex        =   134
      Top             =   4230
      Width           =   195
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   8685
      Picture         =   "frmMain.frx":6DB1A
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   140
      TabIndex        =   98
      Top             =   9000
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11475
      Top             =   4275
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00181818&
      BorderStyle     =   0  'None
      Height          =   75
      Left            =   3420
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   88
      Top             =   7860
      Width           =   750
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   45
      Picture         =   "frmMain.frx":6E54E
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   84
      Top             =   9090
      Width           =   2265
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   225
      TabIndex        =   76
      Top             =   7785
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00AAE0EA&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   270
      TabIndex        =   78
      Top             =   8100
      Width           =   2220
      Begin VB.OptionButton Option2 
         BackColor       =   &H00AAE0EA&
         Caption         =   "-2"
         Height          =   285
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   0
         Width           =   465
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00AAE0EA&
         Caption         =   "-1"
         Height          =   285
         Index           =   1
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   0
         Width           =   465
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00AAE0EA&
         Caption         =   "0"
         Height          =   285
         Index           =   2
         Left            =   900
         MaskColor       =   &H00AAE0EA&
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   0
         Value           =   -1  'True
         Width           =   420
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00AAE0EA&
         Caption         =   "+1"
         Height          =   285
         Index           =   3
         Left            =   1305
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   0
         Width           =   465
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00AAE0EA&
         Caption         =   "+2"
         Height          =   285
         Index           =   4
         Left            =   1755
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":71796
      Height          =   285
      Index           =   15
      Left            =   8265
      Picture         =   "frmMain.frx":71B12
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":71E96
      Height          =   285
      Index           =   14
      Left            =   7725
      Picture         =   "frmMain.frx":72212
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":72596
      Height          =   285
      Index           =   13
      Left            =   7185
      Picture         =   "frmMain.frx":72912
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":72C96
      Height          =   285
      Index           =   12
      Left            =   6645
      Picture         =   "frmMain.frx":73012
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":73396
      Height          =   285
      Index           =   11
      Left            =   6105
      Picture         =   "frmMain.frx":73712
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":73A96
      Height          =   285
      Index           =   10
      Left            =   5565
      Picture         =   "frmMain.frx":73E12
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":74196
      Height          =   285
      Index           =   9
      Left            =   5025
      Picture         =   "frmMain.frx":74512
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":74896
      Height          =   285
      Index           =   8
      Left            =   4485
      Picture         =   "frmMain.frx":74C12
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":74F96
      Height          =   285
      Index           =   7
      Left            =   3945
      Picture         =   "frmMain.frx":75312
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":75696
      Height          =   285
      Index           =   6
      Left            =   3405
      Picture         =   "frmMain.frx":75A12
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":75D96
      Height          =   285
      Index           =   5
      Left            =   2865
      Picture         =   "frmMain.frx":76112
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":76496
      Height          =   285
      Index           =   4
      Left            =   2325
      Picture         =   "frmMain.frx":76812
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":76B96
      Height          =   285
      Index           =   3
      Left            =   1785
      Picture         =   "frmMain.frx":76F12
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":77296
      Height          =   285
      Index           =   2
      Left            =   1245
      Picture         =   "frmMain.frx":77612
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5040
      Width           =   420
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":77996
      Height          =   285
      Index           =   1
      Left            =   705
      Picture         =   "frmMain.frx":77D12
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5040
      Width           =   420
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":78096
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   8250
      Picture         =   "frmMain.frx":78386
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":78676
      Enabled         =   0   'False
      Height          =   285
      Index           =   14
      Left            =   7710
      Picture         =   "frmMain.frx":78966
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":78C56
      Enabled         =   0   'False
      Height          =   285
      Index           =   13
      Left            =   7170
      Picture         =   "frmMain.frx":78F46
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":79236
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   6630
      Picture         =   "frmMain.frx":79526
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":79816
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   6090
      Picture         =   "frmMain.frx":79B06
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":79DF6
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   5550
      Picture         =   "frmMain.frx":7A0E6
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7A3D6
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   5010
      Picture         =   "frmMain.frx":7A6C6
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7A9B6
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   4470
      Picture         =   "frmMain.frx":7ACA6
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7AF96
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   3930
      Picture         =   "frmMain.frx":7B286
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7B576
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   3390
      Picture         =   "frmMain.frx":7B866
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7BB56
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2850
      Picture         =   "frmMain.frx":7BE46
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7C136
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2310
      Picture         =   "frmMain.frx":7C426
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7C716
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1770
      Picture         =   "frmMain.frx":7CA06
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7CCF6
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1230
      Picture         =   "frmMain.frx":7CFE6
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6570
      Width           =   450
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":7D2D6
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   690
      Picture         =   "frmMain.frx":7D5C6
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   6570
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   15
      Left            =   8235
      Picture         =   "frmMain.frx":7D8B6
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   42
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   14
      Left            =   7695
      Picture         =   "frmMain.frx":7E49E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   13
      Left            =   7155
      Picture         =   "frmMain.frx":7F086
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   40
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   12
      Left            =   6615
      Picture         =   "frmMain.frx":7FC6E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   39
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   11
      Left            =   6075
      Picture         =   "frmMain.frx":80856
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   38
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   10
      Left            =   5535
      Picture         =   "frmMain.frx":8143E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   37
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   9
      Left            =   4995
      Picture         =   "frmMain.frx":82026
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   36
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   8
      Left            =   4455
      Picture         =   "frmMain.frx":82C0E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   35
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   7
      Left            =   3915
      Picture         =   "frmMain.frx":837F6
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   34
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   6
      Left            =   3375
      Picture         =   "frmMain.frx":843DE
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   33
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   5
      Left            =   2835
      Picture         =   "frmMain.frx":84FC6
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   32
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   4
      Left            =   2295
      Picture         =   "frmMain.frx":85BAE
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   31
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   3
      Left            =   1755
      Picture         =   "frmMain.frx":86796
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   30
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   2
      Left            =   1215
      Picture         =   "frmMain.frx":8737E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   29
      Top             =   5355
      Width           =   465
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   1
      Left            =   675
      Picture         =   "frmMain.frx":87F66
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   28
      Top             =   5355
      Width           =   465
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00404040&
      DisabledPicture =   "frmMain.frx":88B4E
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   150
      Picture         =   "frmMain.frx":88E3E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6570
      Width           =   450
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      DownPicture     =   "frmMain.frx":8912E
      Height          =   285
      Index           =   0
      Left            =   145
      Picture         =   "frmMain.frx":894AA
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   420
   End
   Begin VB.PictureBox VelocityCtl 
      BackColor       =   &H00848B93&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   9795
      Picture         =   "frmMain.frx":8982E
      ScaleHeight     =   270
      ScaleWidth      =   1275
      TabIndex        =   25
      Top             =   5850
      Width           =   1275
   End
   Begin VB.PictureBox BalanceCtl 
      BackColor       =   &H00848B93&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9795
      Picture         =   "frmMain.frx":8B5A8
      ScaleHeight     =   285
      ScaleWidth      =   1260
      TabIndex        =   24
      Top             =   5475
      Width           =   1260
   End
   Begin VB.PictureBox TransposeCtl 
      BackColor       =   &H00848B93&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9795
      Picture         =   "frmMain.frx":8D322
      ScaleHeight     =   285
      ScaleWidth      =   1260
      TabIndex        =   23
      Top             =   6240
      Width           =   1260
   End
   Begin VB.PictureBox VolumeCtl 
      BackColor       =   &H00848B93&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   9210
      Picture         =   "frmMain.frx":8F09C
      ScaleHeight     =   1065
      ScaleWidth      =   405
      TabIndex        =   19
      Top             =   5475
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E858F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1185
      Index           =   0
      Left            =   135
      Picture         =   "frmMain.frx":9156E
      ScaleHeight     =   1185
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   5355
      Width           =   465
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11430
      Top             =   3825
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   855
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   93
      Top             =   765
      Width           =   9735
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   7650
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   94
         Top             =   4995
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   240
         Left            =   990
         Shape           =   3  'Circle
         Top             =   3555
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3345
      Left            =   855
      Picture         =   "frmMain.frx":92156
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   95
      Top             =   765
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox Picture7 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3075
         Left            =   2475
         Picture         =   "frmMain.frx":B5D8A
         ScaleHeight     =   205
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   123
         TabIndex        =   99
         Top             =   5625
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   2655
         IntegralHeight  =   0   'False
         Left            =   4680
         Sorted          =   -1  'True
         TabIndex        =   97
         Top             =   360
         Width           =   3405
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFF80&
         Height          =   2655
         IntegralHeight  =   0   'False
         Left            =   315
         Sorted          =   -1  'True
         TabIndex        =   96
         Top             =   360
         Width           =   3405
      End
      Begin VB.Image Image10 
         Height          =   3075
         Left            =   2475
         Picture         =   "frmMain.frx":BC136
         Top             =   5625
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Image Image9 
         Height          =   3075
         Left            =   2475
         Picture         =   "frmMain.frx":C28C2
         Top             =   5625
         Visible         =   0   'False
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3345
      Left            =   855
      ScaleHeight     =   3345
      ScaleWidth      =   9735
      TabIndex        =   100
      Top             =   765
      Visible         =   0   'False
      Width           =   9735
      Begin VB.PictureBox Image11 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   960
         Left            =   4815
         ScaleHeight     =   900
         ScaleWidth      =   2790
         TabIndex        =   128
         Top             =   2250
         Width           =   2850
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sa"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   480
            Index           =   7
            Left            =   765
            TabIndex        =   131
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "e"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   480
            Index           =   9
            Left            =   1935
            TabIndex        =   130
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   240
            Left            =   1305
            Shape           =   3  'Circle
            Top             =   90
            Width           =   240
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "mpl"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   480
            Index           =   8
            Left            =   1215
            TabIndex        =   129
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   1935
         Width           =   330
      End
      Begin VB.CommandButton Command9 
         Height          =   285
         Left            =   7020
         Picture         =   "frmMain.frx":C904E
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   495
         Width           =   645
      End
      Begin VB.CommandButton Command8 
         Height          =   285
         Left            =   6345
         Picture         =   "frmMain.frx":C970E
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   495
         Width           =   645
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00000000&
         Height          =   510
         Left            =   8190
         Picture         =   "frmMain.frx":C9DCE
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   2160
         Width           =   1410
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000000&
         Height          =   510
         Left            =   8190
         Picture         =   "frmMain.frx":CA9EE
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   2745
         Width           =   1410
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   2295
         Width           =   330
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   1575
         Width           =   330
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   1215
         Width           =   330
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   855
         Width           =   330
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Caption         =   "..."
         Height          =   195
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   495
         Width           =   330
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Show bouncing ball "
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   225
         TabIndex        =   104
         Top             =   2745
         Value           =   1  'Checked
         Width           =   3390
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Open last list music on startup "
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   225
         TabIndex        =   103
         Top             =   2925
         Width           =   3390
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   127
         Top             =   1935
         Width           =   885
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   126
         Top             =   1935
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7965
         X2              =   7965
         Y1              =   3195
         Y2              =   315
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   5
         Left            =   4815
         TabIndex        =   122
         Top             =   855
         Width           =   2850
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background Image:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   4815
         TabIndex        =   121
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "36"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   4
         Left            =   1800
         TabIndex        =   120
         Top             =   1575
         Width           =   1770
      End
      Begin VB.Label Label16 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   3
         Left            =   1800
         TabIndex        =   112
         Top             =   2295
         Width           =   1725
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Times New Roman"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   111
         Top             =   1215
         Width           =   1770
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   110
         Top             =   855
         Width           =   1725
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   109
         Top             =   495
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   180
         X2              =   7965
         Y1              =   315
         Y2              =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preferences:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   180
         TabIndex        =   108
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Font Size"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   107
         Top             =   1575
         Width           =   1155
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Font Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   106
         Top             =   1215
         Width           =   885
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Ball Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   105
         Top             =   2295
         Width           =   885
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   102
         Top             =   855
         Width           =   885
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   101
         Top             =   495
         Width           =   885
      End
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Channel 10 as Drum"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B9EAEA&
      Height          =   165
      Left            =   5175
      TabIndex        =   135
      Top             =   4230
      Width           =   1305
   End
   Begin VB.Line Line3 
      X1              =   114
      X2              =   165
      Y1              =   534
      Y2              =   534
   End
   Begin VB.Line Line2 
      X1              =   69
      X2              =   18
      Y1              =   534
      Y2              =   534
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Octaves"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1125
      TabIndex        =   133
      Top             =   7920
      Width           =   525
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   8010
      TabIndex        =   132
      Top             =   8010
      Width           =   660
   End
   Begin VB.Image Image8 
      Height          =   270
      Left            =   8685
      Picture         =   "frmMain.frx":CB60E
      Top             =   9000
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image Image7 
      Height          =   270
      Left            =   8685
      Picture         =   "frmMain.frx":CC426
      Top             =   9000
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   9000
      Picture         =   "frmMain.frx":CD23E
      Top             =   9000
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   8865
      Picture         =   "frmMain.frx":CDDFE
      Top             =   9000
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   10800
      TabIndex        =   92
      Top             =   135
      Width           =   465
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   2
      Left            =   5040
      TabIndex        =   91
      Top             =   7830
      Width           =   2805
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   225
      Index           =   1
      Left            =   4455
      TabIndex        =   90
      Top             =   7560
      Width           =   4020
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stopped"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   165
      Left            =   3540
      TabIndex        =   89
      Top             =   7320
      Width           =   540
   End
   Begin VB.Image Image4 
      Height          =   75
      Left            =   1890
      Picture         =   "frmMain.frx":CE9BE
      Top             =   10260
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   3555
      TabIndex        =   87
      Top             =   7530
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   4590
      Picture         =   "frmMain.frx":CEB5A
      Top             =   8310
      Width           =   255
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   225
      Left            =   8100
      TabIndex        =   86
      Top             =   8250
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Left            =   3195
      TabIndex        =   85
      Top             =   8235
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   1260
      Left            =   45
      Picture         =   "frmMain.frx":CF0A2
      Top             =   9090
      Width           =   2265
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   45
      Picture         =   "frmMain.frx":D25BA
      Top             =   9090
      Width           =   2265
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   240
      Left            =   2430
      TabIndex        =   77
      Top             =   7515
      Width           =   285
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   225
      Index           =   0
      Left            =   4455
      TabIndex        =   75
      Top             =   7290
      Width           =   4035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   210
      Index           =   2
      Left            =   9270
      TabIndex        =   74
      Top             =   5085
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   285
      Index           =   1
      Left            =   10755
      TabIndex        =   73
      Top             =   5040
      Width           =   270
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   315
      TabIndex        =   22
      Top             =   7515
      Width           =   2040
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Channel - -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   240
      Left            =   315
      TabIndex        =   21
      Top             =   7200
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lenght"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3195
      TabIndex        =   20
      Top             =   8055
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   8
      Left            =   4590
      TabIndex        =   18
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   9
      Left            =   5130
      TabIndex        =   17
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   10
      Left            =   5670
      TabIndex        =   16
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   11
      Left            =   6210
      TabIndex        =   15
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   12
      Left            =   6750
      TabIndex        =   14
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   13
      Left            =   7290
      TabIndex        =   13
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   14
      Left            =   7830
      TabIndex        =   12
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   15
      Left            =   8370
      TabIndex        =   11
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "08"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   10
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "07"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   6
      Left            =   3510
      TabIndex        =   9
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "06"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   5
      Left            =   2970
      TabIndex        =   8
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "05"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   4
      Left            =   2430
      TabIndex        =   7
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "04"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   3
      Left            =   1890
      TabIndex        =   6
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "03"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   2
      Left            =   1350
      TabIndex        =   5
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "02"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   810
      TabIndex        =   4
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   4740
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Index           =   0
      Left            =   9270
      TabIndex        =   1
      Top             =   4725
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   3960
      TabIndex        =   0
      Top             =   7515
      Width           =   165
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Channel 10 as Drum"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0003405A&
      Height          =   165
      Left            =   5190
      TabIndex        =   136
      Top             =   4245
      Width           =   1305
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Sliders
Dim Sld(15) As New clsSlider
Dim Balance As New ClsSliderHorz
Dim Velocity As New ClsSliderHorz
Dim Transpose As New ClsSliderHorz
Dim Volume As New ClsSliderVert

'DirectX
Dim Dx As New DirectX7
Dim dmSeg As DirectMusicSegment
Dim Seg1 As DirectMusicSegment
Dim dmPerf As DirectMusicPerformance
Dim dmLoader As DirectMusicLoader
Dim dmState As DirectMusicSegmentState
Dim msg As DMUS_NOTIFICATION_PMSG
Dim hEvent As Long
Dim TSig As DMUS_TIMESIGNATURE

'Channel information
Private Type dxCHANNEL
    PatchChanged As Boolean
    VolumeChanged As Boolean
    Instrument As Integer
    Channel As Integer
    Volume As Integer
    Balance As Integer
    Octave As Integer
End Type
Dim ChannelDesc(15) As dxCHANNEL

'Text
Private Type Lyric
    TempoTotal As Double
    TempoAtual As Double
    TxtString As String
    TxtStringLen As Integer
    TextStart As Long
    FraseIndex As Integer
End Type
Dim Lyr() As Lyric
Dim Xini As Long
Dim Phrase() As String

Dim Byt(410) As Byte
Dim MusicPosition As Long, ButtonIndex As Integer, MouseOver As Boolean
Dim LstButtonIndex As Integer, LstMouseOver As Boolean
Dim X1 As Single, ChangePosition As Boolean
Dim DeviceOpen As Boolean, Status As Integer, FileName As String, ActiveChannel As Integer
Dim TxtString As String, Bpm As Single, ppqn, Quarter As Double, BrowseDir As String, Timebase
Dim ListPointer As Integer, ShowBall As Boolean, Crl1 As Long, Crl2 As Long
Dim PanelEnabled As Boolean
Private Sub BrowseForKar()
    
    Dim ndx As Integer
    List1.Clear
    BrowseDir = BrowseForFolder(Me.hWnd, "Browse for kar files")
    If Trim(BrowseDir) = "" Then Exit Sub
    If Right(BrowseDir, 1) <> "\" Then BrowseDir = BrowseDir & "\"

    Files$ = Dir(BrowseDir & "*.kar", vbArchive)
    If Trim(Files$) <> "" Then
        Do
            List1.AddItem BrowseDir & Files$
            Files$ = Dir
            If Trim(Files$) = "" Then Exit Do
        Loop
    End If

End Sub
Sub LastList()
    
    If List2.ListCount = 0 Then
        FileName = ""
        DeviceOpen = False
        Label14 = "0/0"
        Label10(0) = "Play List Not Found"
        Label10(1) = ""
        Label10(2) = ""
        ZeroVariables
        Exit Sub
    End If
    StopMusic
    ListPointer = List2.ListCount - 1
    FileName = List2.List(ListPointer)
    Label14 = Format(ListPointer + 1, "00") & "/" & Format(List2.ListCount, "00")
    ZeroVariables
    GetTitles
    DeviceOpen = False

End Sub
Sub FirstList()
    
    If List2.ListCount = 0 Then
        FileName = ""
        DeviceOpen = False
        Label14 = "0/0"
        Label10(0) = "Play List Not Found"
        Label10(1) = ""
        Label10(2) = ""
        ZeroVariables
        Exit Sub
    End If
    StopMusic
    ListPointer = 0
    FileName = List2.List(0)
    Label14 = Format(ListPointer + 1, "00") & "/" & Format(List2.ListCount, "00")
    ZeroVariables
    GetTitles
    DeviceOpen = False

End Sub
Private Function GetIniFile() As String
    
    Dim IniFile As String
    IniFile = App.Path
    If Right(IniFile, 1) <> "\" Then IniFile = IniFile & "\"
    GetIniFile = IniFile & "Config.cfg"

End Function
Private Sub OpenList()
    
    Filelist = ShowOpen(hWnd, "Play Lists (*.lst)|*.lst", "SoftKar", App.Path)
    If Trim(Filelist) = "" Then Exit Sub
    WriteIni "General", "FileList", CStr(Filelist), GetIniFile
    List2.Clear
    Open Filelist For Input As #1
    Do Until EOF(1)
        Line Input #1, a$
        List2.AddItem a$
    Loop
    Close #1

End Sub
Sub ButtonClick(ByVal Idx As Integer)

    If Not MouseOver Then Exit Sub
    Select Case Idx
        Case 1
            StopMusic
            ListPointer = 0
            Picture5.ZOrder 0
            Picture5.Visible = True
            PanelEnabled = False
        Case 2
            Play
        Case 3
            PauseMusic
        Case 4
            StopMusic
            Image3.Left = 305
        Case 5
            SeekStart
            Image3.Left = 305
            Label4 = "00:00"
            Label8 = "0 %"
        Case 6
            Sec = Int(dmSeg.GetLength / 1000)
            Min = Int(Sec / 60)
            Sec = Int(Sec - (Min * 60))
            Label4 = Format(Min, "00") & ":" & Format(Sec, "00")
            Label8 = "100 %"
            SeekEnd
            Image3.Left = 506
        Case 7
            If Text1.Visible Then
                Text1.Visible = False
            Else
                Text1 = ""
                GetTextLyric
                Text1.Visible = True
            End If
        Case 8
            EndProgram
        Case 9
            PreviousList
        Case 10
            FirstList
        Case 11
            LastList
        Case 12
            NextList
        Case 13
            ShowPreferences
    End Select

End Sub
Sub CloseDevice()
    
    On Error Resume Next
    Timer1.Enabled = False
    If DeviceOpen Then
        dmPerf.Stop dmSeg, dmState, 0, 0
        dmPerf.Stop Seg1, dmState, 0, 0
        dmPerf.CloseDown
        DoEvents
    End If
    dmSeg.Unload dmPerf
    Seg1.Unload dmPerf
    dmPerf.Reset 1
    Set dmPerf = Nothing
    Set dmSeg = Nothing
    Set Seg1 = Nothing
    Set dmState = Nothing
    ZeroVariables
    
End Sub
Sub CreateFileTemp()
    
    'Create temp file
    TempPath = GetTempDir
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
    TempPath = TempPath & "filetemp.tmp"
    Open TempPath For Binary As #2
    Byt(0) = &H4D
    Byt(1) = &H54
    Byt(2) = &H68
    Byt(3) = &H64
    Byt(4) = &H0
    Byt(5) = &H0
    Byt(6) = &H0
    Byt(7) = &H6
    Byt(8) = &H0
    Byt(9) = &H0
    Byt(10) = &H0
    Byt(11) = &H1
    Byt(12) = &H0
    Byt(13) = &H60
            
    Byt(14) = &H4D
    Byt(15) = &H54
    Byt(16) = &H72
    Byt(17) = &H6B
    Byt(18) = &H0
    Byt(19) = &H0
    Byt(20) = &H1
    Byt(21) = &H80
    Byt(22) = &H0
            
    j = 0
    Cn = 23
    Do
        Byt(Cn) = &HC0
        Byt(Cn + 1) = j
        Byt(Cn + 2) = &H1
        j = j + 1
        Cn = Cn + 3
        If j > 127 Then Exit Do
    Loop
    Byt(407) = &H0
    Byt(408) = &HFF
    Byt(409) = &H2F
    Byt(410) = &H0
    Put #2, , Byt()
    Close #2

End Sub
Sub EndProgram()
    
    On Error Resume Next
    Timer1.Enabled = False
    If DeviceOpen Then
        dmPerf.Stop dmSeg, dmState, 0, 0
        dmPerf.CloseDown
        DoEvents
    End If
    TempPath = GetTempDir
    If Right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
    TempPath = TempPath & "filetemp.tmp"
    Kill TempPath
    End
End Sub
Function GetPatch(ByVal P As Integer) As String

    Select Case P + 1
        'PIANO
        Case 1: GetPatch = "Acoustic Grand"
        Case 2: GetPatch = "Bright Acoustic"
        Case 3: GetPatch = "Electric Grand"
        Case 4: GetPatch = "Honky-Tonk"
        Case 5: GetPatch = "Electric Piano 1"
        Case 6: GetPatch = "Electric Piano 2"
        Case 7: GetPatch = "Harpsichord"
        Case 8: GetPatch = "Clavinet"
        'CHROMATIC PERCUSSION
        Case 9: GetPatch = "Celesta"
        Case 10: GetPatch = "Glockenspiel"
        Case 11: GetPatch = "Music Box"
        Case 12: GetPatch = "Vibraphone"
        Case 13: GetPatch = "Marimba"
        Case 14: GetPatch = "Xylophone"
        Case 15: GetPatch = "Tubular Bells"
        Case 16: GetPatch = "Dulcimer"
        'Organ
        Case 17: GetPatch = "Drawbar Organ"
        Case 18: GetPatch = "Percussive Organ"
        Case 19: GetPatch = "Rock Organ"
        Case 20: GetPatch = "Church Organ"
        Case 21: GetPatch = "Reed Organ"
        Case 22: GetPatch = "Accoridan"
        Case 23: GetPatch = "Harmonica"
        Case 24: GetPatch = "Tango Accordian"
        'GUITAR
        Case 25: GetPatch = "Nylon String Guitar"
        Case 26: GetPatch = "Steel String Guitar"
        Case 27: GetPatch = "Electric Jazz Guitar"
        Case 28: GetPatch = "Electric Clean Guitar"
        Case 29: GetPatch = "Electric Muted Guitar"
        Case 30: GetPatch = "Overdriven Guitar"
        Case 31: GetPatch = "Distortion Guitar"
        Case 32: GetPatch = "Guitar Harmonics"
    
        'BASS
        Case 33: GetPatch = "Acoustic Bass"
        Case 34: GetPatch = "Electric Bass(finger)"
        Case 35: GetPatch = "Electric Bass(pick)"
        Case 36: GetPatch = "Fretless Bass"
        Case 37: GetPatch = "Slap Bass 1"
        Case 38: GetPatch = "Slap Bass 2"
        Case 39: GetPatch = "Synth Bass 1"
        Case 40: GetPatch = "Synth Bass 2"
        
        'SOLO STRINGS
        Case 41: GetPatch = "Violin"
        Case 42: GetPatch = "Viola"
        Case 43: GetPatch = "Cello"
        Case 44: GetPatch = "Contrabass"
        Case 45: GetPatch = "Tremolo Strings"
        Case 46: GetPatch = "Pizzicato Strings"
        Case 47: GetPatch = "Orchestral Strings"
        Case 48: GetPatch = "Timpani"
        '  Ensemble
        Case 49: GetPatch = "String Ensemble 1"
        Case 50: GetPatch = "String Ensemble 2"
        Case 51: GetPatch = "SynthStrings 1"
        Case 52: GetPatch = "SynthStrings 2"
        Case 53: GetPatch = "Choir Aahs"
        Case 54: GetPatch = "Voice Oohs"
        Case 55: GetPatch = "Synth Voice"
        Case 56: GetPatch = "Orchestra Hit"
        'BRASS
        Case 57: GetPatch = "Trumpet"
        Case 58: GetPatch = "Trombone"
        Case 59: GetPatch = "Tuba"
        Case 60: GetPatch = "Muted Trumpet"
        Case 61: GetPatch = "French Horn"
        Case 62: GetPatch = "Brass Section"
        Case 63: GetPatch = "SynthBrass 1"
        Case 64: GetPatch = "SynthBrass 2"
        'REED
        Case 65: GetPatch = "Soprano Sax"
        Case 66: GetPatch = "Alto Sax"
        Case 67: GetPatch = "Tenor Sax"
        Case 68: GetPatch = "Baritone Sax"
        Case 69: GetPatch = "Oboe"
        Case 70: GetPatch = "English Horn"
        Case 71: GetPatch = "Bassoon"
        Case 72: GetPatch = "Clarinet"
        'PIPE
        Case 73: GetPatch = "Piccolo"
        Case 74: GetPatch = "Flute"
        Case 75: GetPatch = "Recorder"
        Case 76: GetPatch = "Pan Flute"
        Case 77: GetPatch = "Blown Bottle"
        Case 78: GetPatch = "Skakuhachi"
        Case 79:  GetPatch = "Whistle"
        Case 80: GetPatch = "Ocarina"
        'SYNTH LEAD
        Case 81: GetPatch = "Lead 1 (square)"
        Case 82: GetPatch = "Lead 2 (sawtooth)"
        Case 83: GetPatch = "Lead 3 (calliope)"
        Case 84: GetPatch = "Lead 4 (chiff)"
        Case 85: GetPatch = "Lead 5 (charang)"
        Case 86: GetPatch = "Lead 6 (voice)"
        Case 87: GetPatch = "Lead 7 (fifths)"
        Case 88: GetPatch = "Lead 8 (bass+lead)"
        'SYNTH PAD
        Case 89: GetPatch = "Pad 1 (new age)"
        Case 90: GetPatch = "Pad 2 (warm)"
        Case 91: GetPatch = "Pad 3 (polysynth)"
        Case 92: GetPatch = "Pad 4 (choir)"
        Case 93: GetPatch = "Pad 5 (bowed)"
        Case 94: GetPatch = "Pad 6 (metallic)"
        Case 95: GetPatch = "Pad 7 (halo)"
        Case 96: GetPatch = "Pad 8 (sweep)"
        'SYNTH EFFECTS
        Case 97:  GetPatch = "FX 1 (rain)"
        Case 98:  GetPatch = "FX 2 (soundtrack)"
        Case 99:  GetPatch = "FX 3 (crystal)"
        Case 100: GetPatch = "FX 4 (atmosphere)"
        Case 101: GetPatch = "FX 5 (brightness)"
        Case 102: GetPatch = "FX 6 (goblins)"
        Case 103: GetPatch = "FX 7 (echoes)"
        Case 104: GetPatch = "FX 8 (sci-fi)"
        'ETHNIC
        Case 105: GetPatch = "Sitar"
        Case 106: GetPatch = "Banjo"
        Case 107: GetPatch = "Shamisen"
        Case 108: GetPatch = "Koto"
        Case 109: GetPatch = "Kalimba"
        Case 110: GetPatch = "Bagpipe"
        Case 111: GetPatch = "Fiddle"
        Case 112: GetPatch = "Shanai"
        'PERCUSSIVE
        Case 113: GetPatch = "Tinkle Bell"
        Case 114: GetPatch = "Agogo"
        Case 115: GetPatch = "Steel Drums"
        Case 116: GetPatch = "Woodblock"
        Case 117: GetPatch = "Taiko Drum"
        Case 118: GetPatch = "Melodic Tom"
        Case 119: GetPatch = "Synth Drum"
        Case 120: GetPatch = "Reverse Cymbal"
        'SOUND EFFECTS
        Case 121: GetPatch = "Guitar Fret Noise"
        Case 122: GetPatch = "Breath Noise"
        Case 123: GetPatch = "Seashore"
        Case 124: GetPatch = "Bird Tweet"
        Case 125: GetPatch = "Telephone Ring"
        Case 126: GetPatch = "Helicopter"
        Case 127: GetPatch = "Applause"
        Case 128: GetPatch = "Gunshot"
            
        'Generic channel patch
        Case 129: GetPatch = "Percussion"
    
    End Select

End Function
Private Function GetTip(ByVal nx As Integer) As String
    
    Select Case nx
        Case 1
            GetTip = "Open Play List"
        Case 2
            GetTip = "Play"
        Case 3
            GetTip = "Pause"
        Case 4
            GetTip = "Stop"
        Case 5
            GetTip = "Seek Start"
        Case 6
            GetTip = "Seek End"
        Case 7
            GetTip = "Music Text"
        Case 8
            GetTip = "Exit"
        Case Else
            GetTip = ""
    End Select
    
End Function
Sub GetTitles()
    
    If FileName = "" Then
        Label10(0) = ""
        Label10(1) = ""
        Label10(2) = ""
        Exit Sub
    End If
    
    Dim Titles(0 To 2) As String
    Titles(0) = ""
    Titles(1) = ""
    Titles(2) = ""
    Open FileName For Binary As #1
    TxtString = Space(LOF(1))
    Get #1, , TxtString
    On Error Resume Next
    N = 21
    j = 0
    For t = 1 To 3
    K = InStr(N, TxtString, "@T")
    If K > 0 Then
        If Mid(TxtString, K - 3, 2) = Chr(&HFF) & Chr(&H1) Then
            TextLenght = Asc(Mid(TxtString, K - 1, 1))
            Titles(j) = StrConv(Mid(TxtString, K + 2, TextLenght - 2), vbProperCase)
            j = j + 1
        End If
        N = K + 1
    End If
    Next
    Close #1
    Label10(0) = ""
    Label10(1) = ""
    Label10(2) = ""
    For i = 0 To 2
        Label10(i) = Titles(i)
    Next
    If Trim(Label10(0)) = "" Then Label10(0) = "No titles available"
    Picture4.Cls
    Text1.Visible = False

End Sub
Sub GetWords(ByVal MidString As String)

    On Error Resume Next
    Dim ndx As Long, K As Long, TimeTotal As Double, Pausa As Double, NewPhrase As Integer, Ts As Integer
    Dim Byte1 As Byte, Byte2 As Byte, Byte3 As Byte, Byte4 As Byte
    K = 1: ndx = 0: NewPhrase = 0
    Do
        NewText = ""
        K = InStr(K, MidString, Chr(&HFF))
        If K = 0 Then Exit Sub
        TextLenght = Asc(Mid(MidString, K + 2, 1))
        If Asc(Mid(MidString, K + 1, 1)) = &H1 Then
            If Mid(MidString, K + 3, 1) <> "@" Then
                NewText = Mid(MidString, K + 3, TextLenght)
            End If
        End If
        Byte1 = Asc(Mid(MidString, (K + 3) + TextLenght, 1))
        Byte2 = Asc(Mid(MidString, (K + 4) + TextLenght, 1))
        Byte3 = Asc(Mid(MidString, (K + 5) + TextLenght, 1))
        Byte4 = Asc(Mid(MidString, (K + 6) + TextLenght, 1))
        TempValue = 0
        If Byte2 < &HFF Then
            If Byte3 < &HFF Then
                If Byte4 < &HFF Then
                    TempValue = TempValue And &H7F
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte1 And &H7F)
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte2 And &H7F)
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte3 And &H7F)
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte4 And &H7F)
                    Pausa = TempValue * Quarter
                Else
                    TempValue = TempValue And &H7F
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte1 And &H7F)
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte2 And &H7F)
                    TempValue = TempValue * &H80
                    TempValue = TempValue Or (Byte3 And &H7F)
                    Pausa = TempValue * Quarter
                End If
            Else
                TempValue = (((Byte1 And &H7F) * &H80) Or (Byte2 And &H7F))
                Pausa = TempValue * Quarter
            End If
        Else
            Pausa = Byte1 * Quarter
        End If
        TimeTotal = TimeTotal + Pausa
        ndx = ndx + 1
        
        ReDim Preserve Lyr(ndx)
        Select Case Trim(NewText)
            Case "\", "/"
                Ts = 0
                NewText = ""
                NewPhrase = NewPhrase + 1
                ReDim Preserve Phrase(NewPhrase)
                Lyr(ndx).TxtString = ""
                Lyr(ndx).TxtStringLen = 0
                Lyr(ndx).TextStart = Ts
                Lyr(ndx).TempoAtual = Pausa
                Lyr(ndx).TempoTotal = TimeTotal
                Lyr(ndx).FraseIndex = NewPhrase
                Phrase(NewPhrase) = ""
            Case Else
                If Left(NewText, 1) = "/" Or Left(NewText, 1) = "\" Then
                    Ts = 0
                    NewText = LTrim(Right(NewText, Len(NewText) - 1))
                    NewPhrase = NewPhrase + 1
                    ReDim Preserve Phrase(NewPhrase)
                    Lyr(ndx).TxtString = NewText
                    Lyr(ndx).TxtStringLen = Len(NewText)
                    Lyr(ndx).TextStart = Ts
                    Lyr(ndx).TempoAtual = Pausa
                    Lyr(ndx).TempoTotal = TimeTotal
                    Lyr(ndx).FraseIndex = NewPhrase
                    Phrase(NewPhrase) = NewText
                Else
                    Lyr(ndx).TxtString = NewText
                    Lyr(ndx).TxtStringLen = Len(NewText)
                    Lyr(ndx).TextStart = Ts
                    Lyr(ndx).TempoAtual = Pausa
                    Lyr(ndx).TempoTotal = TimeTotal
                    Lyr(ndx).FraseIndex = NewPhrase
                    Phrase(NewPhrase) = Phrase(NewPhrase) & NewText
                End If
        End Select
        Ts = Ts + Len(NewText)
        K = K + 1
    Loop

End Sub
Sub Initialize()

    Set dmPerf = Dx.DirectMusicPerformanceCreate
    Call dmPerf.Init(Nothing, hWnd)
    dmPerf.SetPort -1, 1
    dmPerf.AddNotificationType DMUS_NOTIFY_ON_MEASUREANDBEAT

End Sub
Sub LoadPatches()
        
        List3.AddItem "Acoustic Grand"
        List3.AddItem "Bright Acoustic"
        List3.AddItem "Electric Grand"
        List3.AddItem "Honky-Tonk"
        List3.AddItem "Electric Piano 1"
        List3.AddItem "Electric Piano 2"
        List3.AddItem "Harpsichord"
        List3.AddItem "Clavinet"
        'CHROMATIC PERCUSSION
        List3.AddItem "Celesta"
        List3.AddItem "Glockenspiel"
        List3.AddItem "Music Box"
        List3.AddItem "Vibraphone"
        List3.AddItem "Marimba"
        List3.AddItem "Xylophone"
        List3.AddItem "Tubular Bells"
        List3.AddItem "Dulcimer"
        'Organ
        List3.AddItem "Drawbar Organ"
        List3.AddItem "Percussive Organ"
        List3.AddItem "Rock Organ"
        List3.AddItem "Church Organ"
        List3.AddItem "Reed Organ"
        List3.AddItem "Accoridan"
        List3.AddItem "Harmonica"
        List3.AddItem "Tango Accordian"
        'GUITAR
        List3.AddItem "Nylon String Guitar"
        List3.AddItem "Steel String Guitar"
        List3.AddItem "Electric Jazz Guitar"
        List3.AddItem "Electric Clean Guitar"
        List3.AddItem "Electric Muted Guitar"
        List3.AddItem "Overdriven Guitar"
        List3.AddItem "Distortion Guitar"
        List3.AddItem "Guitar Harmonics"
    
        'BASS
        List3.AddItem "Acoustic Bass"
        List3.AddItem "Electric Bass(finger)"
        List3.AddItem "Electric Bass(pick)"
        List3.AddItem "Fretless Bass"
        List3.AddItem "Slap Bass 1"
        List3.AddItem "Slap Bass 2"
        List3.AddItem "Synth Bass 1"
        List3.AddItem "Synth Bass 2"
        
        'SOLO STRINGS
        List3.AddItem "Violin"
        List3.AddItem "Viola"
        List3.AddItem "Cello"
        List3.AddItem "Contrabass"
        List3.AddItem "Tremolo Strings"
        List3.AddItem "Pizzicato Strings"
        List3.AddItem "Orchestral Strings"
        List3.AddItem "Timpani"
        '  Ensemble
        List3.AddItem "String Ensemble 1"
        List3.AddItem "String Ensemble 2"
        List3.AddItem "SynthStrings 1"
        List3.AddItem "SynthStrings 2"
        List3.AddItem "Choir Aahs"
        List3.AddItem "Voice Oohs"
        List3.AddItem "Synth Voice"
        List3.AddItem "Orchestra Hit"
        'BRASS
        List3.AddItem "Trumpet"
        List3.AddItem "Trombone"
        List3.AddItem "Tuba"
        List3.AddItem "Muted Trumpet"
        List3.AddItem "French Horn"
        List3.AddItem "Brass Section"
        List3.AddItem "SynthBrass 1"
        List3.AddItem "SynthBrass 2"
        'REED
        List3.AddItem "Soprano Sax"
        List3.AddItem "Alto Sax"
        List3.AddItem "Tenor Sax"
        List3.AddItem "Baritone Sax"
        List3.AddItem "Oboe"
        List3.AddItem "English Horn"
        List3.AddItem "Bassoon"
        List3.AddItem "Clarinet"
        'PIPE
        List3.AddItem "Piccolo"
        List3.AddItem "Flute"
        List3.AddItem "Recorder"
        List3.AddItem "Pan Flute"
        List3.AddItem "Blown Bottle"
        List3.AddItem "Skakuhachi"
        List3.AddItem "Whistle"
        List3.AddItem "Ocarina"
        'SYNTH LEAD
        List3.AddItem "Lead 1 (square)"
        List3.AddItem "Lead 2 (sawtooth)"
        List3.AddItem "Lead 3 (calliope)"
        List3.AddItem "Lead 4 (chiff)"
        List3.AddItem "Lead 5 (charang)"
        List3.AddItem "Lead 6 (voice)"
        List3.AddItem "Lead 7 (fifths)"
        List3.AddItem "Lead 8 (bass+lead)"
        'SYNTH PAD
        List3.AddItem "Pad 1 (new age)"
        List3.AddItem "Pad 2 (warm)"
        List3.AddItem "Pad 3 (polysynth)"
        List3.AddItem "Pad 4 (choir)"
        List3.AddItem "Pad 5 (bowed)"
        List3.AddItem "Pad 6 (metallic)"
        List3.AddItem "Pad 7 (halo)"
        List3.AddItem "Pad 8 (sweep)"
        'SYNTH EFFECTS
        List3.AddItem "FX 1 (rain)"
        List3.AddItem "FX 2 (soundtrack)"
        List3.AddItem "FX 3 (crystal)"
        List3.AddItem "FX 4 (atmosphere)"
        List3.AddItem "FX 5 (brightness)"
        List3.AddItem "FX 6 (goblins)"
        List3.AddItem "FX 7 (echoes)"
        List3.AddItem "FX 8 (sci-fi)"
        'ETHNIC
        List3.AddItem "Sitar"
        List3.AddItem "Banjo"
        List3.AddItem "Shamisen"
        List3.AddItem "Koto"
        List3.AddItem "Kalimba"
        List3.AddItem "Bagpipe"
        List3.AddItem "Fiddle"
        List3.AddItem "Shanai"
        'PERCUSSIVE
        List3.AddItem "Tinkle Bell"
        List3.AddItem "Agogo"
        List3.AddItem "Steel Drums"
        List3.AddItem "Woodblock"
        List3.AddItem "Taiko Drum"
        List3.AddItem "Melodic Tom"
        List3.AddItem "Synth Drum"
        List3.AddItem "Reverse Cymbal"
        'SOUND EFFECTS
        List3.AddItem "Guitar Fret Noise"
        List3.AddItem "Breath Noise"
        List3.AddItem "Seashore"
        List3.AddItem "Bird Tweet"
        List3.AddItem "Telephone Ring"
        List3.AddItem "Helicopter"
        List3.AddItem "Applause"
        List3.AddItem "Gunshot"

End Sub
Sub OpenMusic()
    
    On Error Resume Next
    MousePointer = 11
    CloseDevice
    ReadFile
    Initialize
    DoEvents
    TempFile = GetTempDir()
    If Right(TempFile, 1) <> "\" Then TempFile = TempFile & "\"
    TempFile = TempFile & "filetemp.tmp"
    Set dmSeg = dmLoader.LoadSegment(FileName)
    Set Seg1 = dmLoader.LoadSegment(TempFile)
    
    dmSeg.SetStandardMidiFile
    Seg1.SetStandardMidiFile
    dmSeg.Download dmPerf
    Seg1.Download dmPerf
    
    Min = Format(Int(dmSeg.GetLength / 60000), "00")
    Sec = Format(Int((dmSeg.GetLength / 1000) - (Min * 60)), "00")
    Label5 = "Lenght  " & Min & ":" & Sec
    
    mtTime = dmPerf.GetMusicTime()
    Call dmPerf.PlaySegment(dmSeg, 0, mtTime + 2000)
    Call dmPerf.GetTimeSig(mtTime + 2000, 0, TSig)
    Label9 = TSig.beatsPerMeasure & "/" & TSig.beat
    Call dmPerf.Stop(dmSeg, Nothing, 0, 0)
    
    dmPerf.SetMasterTempo 1
    Status = 1
    MusicPosition = 0
    Volume.Value = 64
    Balance.Value = 64
    Velocity.Value = 64
    Transpose.Value = 64
    dmPerf.SetMasterVolume ((64 * 42) - 3000)
    Xini = 0
    Timer2.Interval = Int(ppqn / (768000 / Timebase))
    Shape1.Top = 48
    DeviceOpen = True
    MousePointer = 0
   
End Sub
Sub PauseMusic()

    On Error Resume Next
    MusicPosition = dmState.GetSeek - 100
    dmPerf.Stop dmSeg, dmState, 0, 0
    Status = 2

End Sub
Sub Play()
    
    On Error Resume Next
    MousePointer = 11
    If Not DeviceOpen Then OpenMusic
    s = (dmPerf.GetMasterTempo * 1000) - Bpm
    NewPosition = (((MusicPosition / 768000) * ppqn) - s)
    Xini = UBound(Lyr)
    For j = 0 To UBound(Lyr) - 1
        If Lyr(j).TempoTotal > NewPosition Then
            Xini = j
            Exit For
        End If
    Next

    If MusicPosition >= dmSeg.GetLength Then MusicPosition = MusicPosition - 10
    If MusicPosition < 0 Then MusicPosition = 0
    dmSeg.SetStartPoint MusicPosition
    Status = 0
    Set dmState = dmPerf.PlaySegment(dmSeg, 0, 0)
    SetChannels
    Timer1.Enabled = True
    Image3.Enabled = True
    MousePointer = 0
    

End Sub
Sub ReadFile()

    On Error Resume Next
    Open FileName For Binary As #1
    TxtString = Space(LOF(1))
    Get #1, , TxtString
    Close #1
    
    Dim isLyric As Boolean
    isLyric = False
    MidiType = Val("&H" & Hex(Asc(Mid(TxtString, 9, 1))) & Hex(Asc(Mid(TxtString, 10, 1))))
    NrTracks = Val("&H" & Hex(Asc(Mid(TxtString, 11, 1))) & Hex(Asc(Mid(TxtString, 12, 1))))
    
    K = InStr(1, TxtString, Chr(&HFF) & Chr(&H51))
    If K <> 0 Then
        n1 = Hex(Asc(Mid(TxtString, K + 3, 1)))
        n2 = Hex(Asc(Mid(TxtString, K + 4, 1)))
        N3 = Hex(Asc(Mid(TxtString, K + 5, 1)))
        If Len(n1) = 1 Then n1 = "0" & n1
        If Len(n2) = 1 Then n2 = "0" & n2
        If Len(N3) = 1 Then N3 = "0" & N3
        T1 = Hex(Asc(Mid(TxtString, 13, 1)))
        T2 = Hex(Asc(Mid(TxtString, 14, 1)))
        If Len(T1) = 1 Then T1 = "0" & T1
        If Len(T2) = 1 Then T2 = "0" & T2
        Timebase = CDec("&H" & T1 & T2) / 4
        Bpm = Format(60000000 / CDec("&H" & n1 & n2 & N3), "0.00")
        ppqn = CDec("&H" & n1 & n2 & N3)
        Quarter = (ppqn / Timebase) / 4000
    End If
    
    IniTrack = 15
    For i = 1 To NrTracks
        B1 = Hex(Asc(Mid(TxtString, IniTrack + 4, 1)))
        B2 = Hex(Asc(Mid(TxtString, IniTrack + 5, 1)))
        B3 = Hex(Asc(Mid(TxtString, IniTrack + 6, 1)))
        B4 = Hex(Asc(Mid(TxtString, IniTrack + 7, 1)))
        If Len(B1) = 1 Then B1 = "0" & B1
        If Len(B2) = 1 Then B2 = "0" & B2
        If Len(B3) = 1 Then B3 = "0" & B3
        If Len(B4) = 1 Then B4 = "0" & B4
        TrackLenght = CLng("&H" & B1 & B2 & B3 & B4)
        TempString = Mid(TxtString, IniTrack + 8, TrackLenght)
        K = InStr(1, TempString, Chr(&HFF) & Chr(&H3) & Chr(&H5) & Chr(&H57) & Chr(&H6F) & Chr(&H72) & Chr(&H64) & Chr(&H73))
        If K > 0 Then
            GetWords TempString
            isLyric = True
            GoTo L1
        Else
            K = 1
            isLyric = True
            For Y = 1 To 50
                K = InStr(K, TempString, Chr(&HFF) & Chr(&H1))
                If K = 0 Then
                    isLyric = False
                    Exit For
                End If
                K = K + 1
            Next
            If isLyric Then
                GetWords TempString
                GoTo L1
            End If
        End If
        IniTrack = IniTrack + TrackLenght + 8
    Next
    GetWords TxtString
    
L1:
    
    MaxWidth = 0
    tempVar = 0
    Cr = ReadIni("General", "MaxFontSize", GetIniFile)
    If Not IsNumeric(Cr) Then Mx = 24 Else Mx = CSng(Cr)
    Picture4.FontSize = Mx
    For i = 0 To UBound(Phrase)
        If tempVar < Picture4.TextWidth(Phrase(i)) Then
            tempVar = Picture4.TextWidth(Phrase(i))
            MaxWidth = i
        End If
    Next
    Do
        If Picture4.TextWidth(Phrase(MaxWidth)) < (Picture4.Width - 10) Then Exit Do
        Picture4.FontSize = Picture4.FontSize - 1
    Loop


    'get instruments descriptions
    For t = 0 To 15
        X1 = 21
        X2 = 21
        X3 = 21
        ChannelDesc(t).Channel = -1
        X0 = InStr(1, TxtString, Chr(&H90 + t))
        If X0 > 0 Then
            ChannelDesc(t).Channel = t + 1
            Do
                X1 = InStr(X1 + 1, TxtString, Chr(&HC0 + t))
                If X1 = 0 Then Exit Do
                po = Asc(Mid(TxtString, X1 + 1, 1))
                If Asc(Mid(TxtString, X1 + 1, 1)) > 0 Then
                    ChannelDesc(t).Instrument = Asc(Mid(TxtString, X1 + 1, 1))
                    Exit Do
                End If
            Loop
            Do
                ChannelDesc(t).Volume = 100
                X2 = InStr(X2 + 1, TxtString, Chr(&HB0 + t) & Chr(7))
                If X2 = 0 Then Exit Do
                If Asc(Mid(TxtString, X2 + 2, 1)) > 0 And Asc(Mid(TxtString, X2 + 2, 1)) < 127 Then
                    ChannelDesc(t).Volume = Asc(Mid(TxtString, X2 + 2, 1))
                    Exit Do
                End If
            Loop
            X3 = InStr(X3 + 1, TxtString, Chr(&HB0 + t) & Chr(10))
            If X3 > 0 Then
                ChannelDesc(t).Balance = Asc(Mid(TxtString, X3 + 2, 1))
                If Asc(Mid(TxtString, X3 + 2, 1)) > 127 Or Asc(Mid(TxtString, X3 + 2, 1)) <= 0 Then
                    X3 = InStr(X3 + 1, TxtString, Chr(&HB0 + t) & Chr(10))
                    If X3 > 0 Then
                        ChannelDesc(t).Balance = Asc(Mid(TxtString, X3 + 2, 1))
                    End If
                End If
            End If
        End If
    Next

    Check4.Enabled = False
    For t = 0 To 15
        ChannelDesc(t).PatchChanged = False
        ChannelDesc(t).VolumeChanged = False
        If ChannelDesc(t).Channel >= 0 Then
            Sld(t).Volume = ChannelDesc(t).Volume
            Sld(t).Balance = ChannelDesc(t).Balance
            Sld(t).Enabled = True
            ChannelDesc(t).Octave = 2
            Label3(t).ForeColor = &HFF00&
            Check1(t).Value = 1
            Option1(t).Enabled = True
            Option1(t).Value = False
            If t = 9 Then
                Check4 = 1
                Check4.Enabled = True
            End If
        Else
            Sld(t).Enabled = False
            Label3(t).ForeColor = &H8000&
            Option1(t).Enabled = False
            Option1(t).Value = False
        End If
        DoEvents
    Next
    CreateFileTemp

End Sub
Sub SeekEnd()
    
    On Error Resume Next
    MusicPosition = dmSeg.GetLength - 250
    dmPerf.Stop Nothing, Nothing, 0, 0
    Picture4.Cls
    Xini = UBound(Lyr) - 10
    Status = 2
    
End Sub
Sub SeekStart()
    
    On Error Resume Next
    StopMusic

End Sub
Sub SetChannels()
    
    On Error Resume Next
    For i = 0 To 15
        If ChannelDesc(i).Channel <> -1 Then
            If ChannelDesc(i).PatchChanged Then
                dmPerf.SendPatchPMSG 0, DMUS_PMSGF_REFTIME, i, ChannelDesc(i).Instrument, 0, 0
            End If
        End If
    Next
    For i = 0 To 15
        If ChannelDesc(i).Channel <> -1 Then
            If ChannelDesc(i).VolumeChanged Then
                dmPerf.SendMIDIPMSG 10, DMUS_PMSGF_REFTIME, i, &HB0, &H7, Sld(i).Volume
                dmPerf.SendMIDIPMSG 20, DMUS_PMSGF_REFTIME, i, &HB0, &HA, Sld(i).Balance
            End If
        End If
    Next

End Sub
Sub SetMusicPosition()

    On Error Resume Next
    If Not DeviceOpen Then Exit Sub
    X = Image3.Left
    MusicPosition = ((X - 304) * 100) / (220 - Image3.Width)
    MusicPosition = (Int(MusicPosition * dmSeg.GetLength) / 100) - 100
    Status = 2
    Play
    
End Sub
Sub SetPreferences()

    Dim Mx As Single
    Dim Crl As Long
    
    'BallColor
    Cr = ReadIni("General", "BallColor", GetIniFile)
    If IsNumeric(Cr) Then Shape1.FillColor = CLng(Cr)
    
    ShowBall = Val(ReadIni("General", "ShowBall", GetIniFile))
    
    'textColor
    Cr = ReadIni("General", "TextColor1", GetIniFile)
    If IsNumeric(Cr) Then Crl1 = CLng(Cr) Else Crl1 = &HFF
    
    'TextColor
    Cr = ReadIni("General", "TextColor2", GetIniFile)
    If IsNumeric(Cr) Then Crl2 = CLng(Cr) Else Crl2 = &HFFFF&

    'Set Font
    FtName = ReadIni("General", "FontName", GetIniFile)
    Ftbold = ReadIni("General", "FontBold", GetIniFile)
    FtItalic = ReadIni("General", "FontItalic", GetIniFile)
    If Trim(FtName) <> "" Then
        Picture4.FontName = FtName
        Picture4.FontBold = Ftbold
        Picture4.FontItalic = FtItalic
    Else
        Picture4.FontName = "Times New Roman"
        Picture4.FontBold = True
        Picture4.FontItalic = False
    End If
    Cr = ReadIni("General", "MaxFontSize", GetIniFile)
    If Not IsNumeric(Cr) Then Mx = 36 Else Mx = CSng(Cr)
    Picture4.FontSize = Mx

    'Picture BackColor
    Cr = ReadIni("General", "BackColor", GetIniFile)
    If IsNumeric(Cr) Then Crl = CLng(Cr) Else Crl = &H0
    Picture4.BackColor = Crl

    'Last PlayList
    If Val(ReadIni("General", "LastList", GetIniFile)) = 1 Then
        Filelist = ReadIni("General", "FileList", GetIniFile)
        If Trim(Filelist) <> "" Then
            If Dir(Filelist, vbArchive) <> "" Then
                List2.Clear
                Open Filelist For Input As #1
                Do Until EOF(1)
                    Line Input #1, a$
                    List2.AddItem a$
                Loop
                Close #1
                If Not DeviceOpen Then
                    If ListPointer <= 0 Then FirstList
                End If
            End If
        End If
    End If

    'Background Image
    ImgStr = ReadIni("General", "ImgBkg", GetIniFile)
    If Trim(ImgStr) = "" Then ImgStr = "None"
    If ImgStr <> "None" Then
        If Dir(ImgStr, vbArchive) <> "" Then
            Picture4.PaintPicture LoadPicture(ImgStr), 0, 0, Picture4.Width, Picture4.Height, 0, 0
            Picture4.Picture = Picture4.Image
        Else
            Picture4 = LoadPicture
        End If
    Else
        Picture4 = LoadPicture
    End If
    PanelEnabled = True


End Sub
Sub SetSample()
    
    Label15(7).ForeColor = Label16(0).BackColor
    Label15(9).ForeColor = Label16(0).BackColor
    Label15(8).ForeColor = Label16(1).BackColor
    Set Label15(7).Font = Picture8.Font
    Set Label15(8).Font = Picture8.Font
    Set Label15(9).Font = Picture8.Font
    Shape2.FillColor = Label16(3).BackColor
    
    L = ((Image11.Width / 2) - (Picture8.TextWidth("Sample") / 2))
    Label15(7).Left = L
    Label15(8).Left = Label15(7).Left + Label15(7).Width
    Label15(9).Left = Label15(8).Left + Label15(8).Width
    Label15(8).ZOrder 0
    Image11.BackColor = Label16(6).BackColor
    If Label16(5) <> "None" Then
        Image11.PaintPicture LoadPicture(Label16(5)), 0, 0, Image11.Width, Image11.Height, 0, 0
    End If

End Sub
Sub ShowPreferences()
    
    On Error Resume Next
    
    If DeviceOpen Then StopMusic
    Label16(0).BackColor = CLng(ReadIni("General", "TextColor1", GetIniFile))
    Label16(1).BackColor = CLng(ReadIni("General", "TextColor2", GetIniFile))
    FtName = ReadIni("General", "FontName", GetIniFile)
    Ftbold = ReadIni("General", "FontBold", GetIniFile)
    FtItalic = ReadIni("General", "FontItalic", GetIniFile)
    If Trim(FtName) <> "" Then
        Label16(2) = FtName
        Picture8.FontName = FtName
        Picture8.FontBold = Ftbold
        Picture8.FontItalic = FtItalic
        Picture8.FontSize = 22
    End If
    Label16(4) = CSng(ReadIni("General", "MaxFontSize", GetIniFile))
    Label16(3).BackColor = CLng(ReadIni("General", "BallColor", GetIniFile))
    Check3 = Val(ReadIni("General", "ShowBall", GetIniFile))
    Check2 = Val(ReadIni("General", "LastList", GetIniFile))
    Label16(6).BackColor = CLng(ReadIni("General", "BackColor", GetIniFile))
    
    ImgStr = ReadIni("General", "ImgBkg", GetIniFile)
    If Trim(ImgStr) = "" Then Label16(5) = "None" Else Label16(5) = ImgStr
    SetSample
    Picture8.Visible = True
    Picture8.ZOrder 0
    PanelEnabled = False

End Sub
Sub StopMusic()
    
    On Error Resume Next
    Status = 1
    dmPerf.Stop Nothing, Nothing, 0, 0
    MusicPosition = 0
    Picture4.Cls
    
End Sub
Sub UpdateBar(ByVal Vle As Integer)
    
    On Error Resume Next
    Image3.Left = (((dmState.GetSeek * 100) / dmSeg.GetLength) * ((218 - Image3.Width) / 100) + 305)

End Sub
Sub ZeroVariables()

    Xini = 0
    ReDim Phrase(0)
    ReDim Lyr(0)
    For j = 0 To 15
        ChannelDesc(j).Balance = 64
        ChannelDesc(j).Channel = -1
        ChannelDesc(j).Instrument = 0
        ChannelDesc(j).Octave = 0
        ChannelDesc(j).Volume = 0
        ChannelDesc(j).PatchChanged = False
        ChannelDesc(j).VolumeChanged = False
        Sld(j).Volume = 0
        Sld(j).Balance = 64
        Sld(j).Enabled = False
        Label3(j).ForeColor = &H8000&
        Check1(j) = 0
    Next
    Picture4.Cls
    Image3.Enabled = False
    Image3.Left = 306
    Option2(2) = True
    Label7 = ""
    Check4 = 0
    Check4.Enabled = False
End Sub
Private Sub BalanceCtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Balance.DragButton Button, X, Y
    Label2(2) = "Balance"
    If Button = 1 Then
        For t = 0 To 15
            dmPerf.SendMIDIPMSG 0, DMUS_PMSGF_REFTIME, t, &HB0, &HA, Balance.Value
        Next
    End If
    Vlr = Int((30 / 127) * Balance.Value)
    If Vlr <> 15 Then
        Label2(1) = IIf(Vlr < 15, "-" & 15 - Vlr, "+" & Vlr - 15)
    Else
        Label2(1) = "0"
    End If
    
End Sub
Private Sub Check1_Click(Index As Integer)
    
    On Error Resume Next
    If Check1(Index) = 0 Then
        Sld(Index).Volume = 0
        Sld(Index).Enabled = False
    Else
        Sld(Index).Volume = ChannelDesc(Index).Volume
        Sld(Index).Enabled = True
    End If
    If DeviceOpen Then
        dmPerf.SendMIDIPMSG 0, DMUS_PMSGF_REFTIME, Index, &HB0, &H7, Sld(Index).Volume
        ChannelDesc(Index).VolumeChanged = True
    End If

End Sub
Private Sub NextList()
    
    If List2.ListCount = 0 Then
        FileName = ""
        DeviceOpen = False
        Label14 = "0/0"
        Label10(0) = "Play List Not Found"
        Label10(1) = ""
        Label10(2) = ""
        ZeroVariables
        Exit Sub
    End If
    StopMusic
    Select Case ListPointer
        Case Is >= List2.ListCount - 1
            ListPointer = List2.ListCount - 1
        Case Else
            ListPointer = ListPointer + 1
    End Select
    FileName = List2.List(ListPointer)
    Label14 = Format(ListPointer + 1, "00") & "/" & Format(List2.ListCount, "00")
    ZeroVariables
    GetTitles
    DeviceOpen = False
    
End Sub
Private Sub SaveList()

    SaveFile = ShowSave(hWnd, "Playlist1.lst", "Play Lists (*.lst)|*.lst", "SoftKar", App.Path)
    If Trim(SaveFile) = "" Then Exit Sub
    Open SaveFile For Output As #1
    For i = 0 To List2.ListCount - 1
        Print #1, List2.List(i)
    Next
    Close #1

End Sub
Private Sub PreviousList()
    
    If List2.ListCount = 0 Then
        FileName = ""
        DeviceOpen = False
        Label14 = "0/0"
        Label10(0) = "Play List Not Found"
        Label10(1) = ""
        Label10(2) = ""
        ZeroVariables
        Exit Sub
    End If
    StopMusic
    Select Case ListPointer
        Case Is <= 0
            ListPointer = 0
        Case Else
            ListPointer = ListPointer - 1
    End Select
    FileName = List2.List(ListPointer)
    Label14 = Format(ListPointer + 1, "00") & "/" & Format(List2.ListCount, "00")
    ZeroVariables
    GetTitles
    DeviceOpen = False
    
End Sub

Private Sub Check3_Click()

    If Check3 = 1 Then
        Shape2.Visible = True
    Else
        Shape2.Visible = False
    End If

End Sub

Private Sub Command1_Click()

    Dim Crl As Long
    Crl = ShowColor(hWnd)
    If Crl = -1 Then Exit Sub
    Label16(0).BackColor = Crl
    SetSample

End Sub
Private Sub Command10_Click()
    
    Dim Crl As Long
    Crl = ShowColor(hWnd)
    If Crl = -1 Then Exit Sub
    Label16(6).BackColor = Crl
    SetSample

End Sub
Private Sub Command2_Click()
    
    Dim Crl As Long
    Crl = ShowColor(hWnd)
    If Crl = -1 Then Exit Sub
    Label16(1).BackColor = Crl
    SetSample
    
End Sub
Private Sub Command3_Click()

    If ShowFont(Me, Label16(2), , Val(Label16(4))) = 0 Then Exit Sub
    Label16(2) = NewFont.FontName
    Label16(4) = NewFont.FontSize
    Picture8.FontName = NewFont.FontName
    Picture8.FontBold = NewFont.FontBold
    Picture8.FontItalic = NewFont.FontItalic
    Picture8.FontSize = 22
    SetSample

End Sub
Private Sub Command4_Click()
    
    If ShowFont(Me, Label16(2), , Val(Label16(4))) = 0 Then Exit Sub
    Label16(4) = NewFont.FontSize

End Sub
Private Sub Command5_Click()
    
    Dim Crl As Long
    Crl = ShowColor(hWnd)
    If Crl = -1 Then Exit Sub
    Label16(3).BackColor = Crl
    SetSample

End Sub
Private Sub Command6_Click()
    
    WriteIni "General", "TextColor1", CStr(Label16(0).BackColor), GetIniFile
    WriteIni "General", "TextColor2", CStr(Label16(1).BackColor), GetIniFile
    WriteIni "General", "FontName", Label16(2), GetIniFile
    WriteIni "General", "FontBold", Picture8.FontBold, GetIniFile
    WriteIni "General", "FontItalic", Picture8.FontItalic, GetIniFile
    WriteIni "General", "MaxFontSize", Val(Label16(4)), GetIniFile
    WriteIni "General", "BallColor", CStr(Label16(3).BackColor), GetIniFile
    WriteIni "General", "ShowBall", CStr(Check3.Value), GetIniFile
    WriteIni "General", "LastList", CStr(Check2.Value), GetIniFile
    WriteIni "General", "BackColor", CStr(Label16(6).BackColor), GetIniFile
    WriteIni "General", "ImgBkg", Label16(5), GetIniFile
    
    Picture8.ZOrder 1
    Picture8.Visible = False
    SetPreferences
    If DeviceOpen Then OpenMusic

    
End Sub
Private Sub Command7_Click()
    
    Picture8.ZOrder 1
    Picture8.Visible = False
    PanelEnabled = True
    
End Sub
Private Sub Command8_Click()

    FileImage = ShowOpen(hWnd, "Bitmaps (*.bmp)|*.bmp", "SoftKar", App.Path)
    If Trim(FileImage) = "" Then Exit Sub
    Label16(5) = FileImage
    SetSample

End Sub
Private Sub Command9_Click()

    Label16(5) = "None"
    SetSample

End Sub
Private Sub Form_Click()

    List3.Visible = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ButtonIndex = 0
    If Not PanelEnabled Then Exit Sub
    If Button = 1 Then
        'Preference Button
        If (X > 595 And X < 709) And (Y > 277 And Y < 292) Then
            PaintPicture Image5, 593, 278
            ButtonIndex = 13
            Exit Sub
        End If
    
        'List Buttons
        If (X > 604 And X < 745) And (Y > 452 And Y < 471) Then
            Select Case Picture6.Point(X - 605, Y - 453)
                Case RGB(255, 0, 0)
                    PaintPicture Image8, 606, 454, 27, 15, 1, 1, 27, 15
                    ButtonIndex = 9
                Case RGB(0, 255, 0)
                    PaintPicture Image8, 606 + 29, 454, 27, 15, 30, 1, 27, 15
                    ButtonIndex = 10
                Case RGB(0, 0, 255)
                    PaintPicture Image8, 606 + 81, 454, 27, 15, 82, 1, 27, 15
                    ButtonIndex = 11
                Case RGB(255, 0, 255)
                    PaintPicture Image8, 606 + 110, 454, 27, 15, 111, 1, 27, 15
                    ButtonIndex = 12
            End Select
            Exit Sub
        End If
        
        'Master Buttons
        Select Case Picture2.Point(X - 599, Y - 484)
            Case RGB(255, 0, 0)
                PaintPicture Image1, 599 + 4, 484 + 3, 36, 36, 4, 3, 36, 36
                ButtonIndex = 1
            Case RGB(255, 255, 0)
                PaintPicture Image1, 599 + 40, 484 + 3, 36, 36, 40, 3, 36, 36
                ButtonIndex = 2
            Case RGB(0, 255, 0)
                If Not DeviceOpen Then Exit Sub
                PaintPicture Image1, 599 + 76, 484 + 3, 36, 36, 76, 3, 36, 36
                ButtonIndex = 3
            Case RGB(0, 0, 255)
                If Not DeviceOpen Then Exit Sub
                PaintPicture Image1, 599 + 112, 484 + 3, 36, 36, 112, 3, 36, 36
                ButtonIndex = 4
            Case RGB(255, 0, 255)
                If Not DeviceOpen Then Exit Sub
                PaintPicture Image1, 599 + 4, 484 + 44, 36, 36, 4, 44, 36, 36
                ButtonIndex = 5
            Case RGB(255, 255, 255)
                If Not DeviceOpen Then Exit Sub
                PaintPicture Image1, 599 + 40, 484 + 44, 36, 36, 40, 44, 36, 36
                ButtonIndex = 6
            Case RGB(0, 0, 0)
                If Not DeviceOpen Then Exit Sub
                PaintPicture Image1, 599 + 76, 484 + 44, 36, 36, 76, 44, 36, 36
                ButtonIndex = 7
            Case RGB(0, 255, 255)
                PaintPicture Image1, 599 + 112, 484 + 44, 36, 36, 112, 44, 36, 36
                ButtonIndex = 8
        End Select
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If ButtonIndex = 0 Then Exit Sub
    MouseOver = False
    If Button = 1 Then
        'Preference Button
        If ButtonIndex = 13 Then
            If (X > 595 And X < 709) And (Y > 277 And Y < 292) Then MouseOver = True
            PaintPicture Image6, 593, 278
            ButtonClick 13
        End If
    
        'List Buttons
        Select Case ButtonIndex
            Case 9
                PaintPicture Image7, 606, 454, 27, 15, 1, 1, 27, 15
                If Picture6.Point(X - 605, Y - 453) = RGB(255, 0, 0) Then MouseOver = True
                ButtonClick 9
            Case 10
                PaintPicture Image7, 606 + 29, 454, 27, 15, 30, 1, 27, 15
                If Picture6.Point(X - 605, Y - 453) = RGB(0, 255, 0) Then MouseOver = True
                ButtonClick 10
            Case 11
                PaintPicture Image7, 606 + 81, 454, 27, 15, 82, 1, 27, 15
                If Picture6.Point(X - 605, Y - 453) = RGB(0, 0, 255) Then MouseOver = True
                ButtonClick 11
            Case 12
                PaintPicture Image7, 606 + 110, 454, 27, 15, 111, 1, 27, 15
                If Picture6.Point(X - 605, Y - 453) = RGB(255, 0, 255) Then MouseOver = True
                ButtonClick 12
        End Select
    
        'Master Buttons
        Select Case ButtonIndex
            Case 1
                PaintPicture Image2, 599 + 4, 484 + 3, 36, 36, 4, 3, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(255, 0, 0) Then MouseOver = True
                ButtonClick 1
            Case 2
                PaintPicture Image2, 599 + 40, 484 + 3, 36, 36, 40, 3, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(255, 255, 0) Then MouseOver = True
                ButtonClick 2
            Case 3
                PaintPicture Image2, 599 + 76, 484 + 3, 36, 36, 76, 3, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(0, 255, 0) Then MouseOver = True
                ButtonClick 3
            Case 4
                PaintPicture Image2, 599 + 112, 484 + 3, 36, 36, 112, 3, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(0, 0, 255) Then MouseOver = True
                ButtonClick 4
            Case 5
                PaintPicture Image2, 599 + 4, 484 + 44, 36, 36, 4, 44, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(255, 0, 255) Then MouseOver = True
                ButtonClick 5
            Case 6
                PaintPicture Image2, 599 + 40, 484 + 44, 36, 36, 40, 44, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(255, 255, 255) Then MouseOver = True
                ButtonClick 6
            Case 7
                PaintPicture Image2, 599 + 76, 484 + 44, 36, 36, 76, 44, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(0, 0, 0) Then MouseOver = True
                ButtonClick 7
            Case 8
                PaintPicture Image2, 599 + 112, 484 + 44, 36, 36, 112, 44, 36, 36
                If Picture2.Point(X - 599, Y - 484) = RGB(0, 255, 255) Then MouseOver = True
                ButtonClick 8
        End Select
    End If

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    ChangePosition = True
    dmSeg.SetStartPoint 0
    dmPerf.Stop Nothing, Nothing, 0, 0
    X1 = Int(ScaleX(X, 1, 3))
    Picture4.Cls
    Status = 2

End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not DeviceOpen Then Exit Sub
    Form_MouseMove Button, Shift, Int(ScaleX(X, 1, 3)) + Image3.Left, ScaleY(Y, 1, 3) + Image3.Top

End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ChangePosition = False
    SetMusicPosition
    Status = 0

End Sub

Private Sub Label1_Change()
    
    Picture3.Cls
    If Status = 0 Then
        w = Picture3.Width / TSig.beatsPerMeasure
        Picture3.PaintPicture Image4, (Picture3.Width / 2) - ((w * (msg.lField1 + 1)) / 2), 0, (w * (msg.lField1 + 1)), 4
    End If

End Sub

Private Sub Label12_Change()
    
    Select Case Label12
        Case "Playing"
            If Xini + 1 > UBound(Lyr) Then
                Timer2.Enabled = False
                Shape1.Visible = False
                Picture4.Cls
            End If
        Case "Stopped"
            Timer2.Enabled = False
            Shape1.Visible = False
            Picture3.Cls
            Timer1.Enabled = False
        Case "Paused"
            Timer2.Enabled = False
            Shape1.Visible = False
            Picture3.Cls
            Timer1.Enabled = False
    End Select

End Sub

Private Sub Label13_Click()

    Me.WindowState = 1

End Sub

Private Sub Label18_Click()

    If Check4.Enabled Then
        If Check4 = 0 Then Check4 = 1 Else Check4 = 0
    End If

End Sub

Private Sub List1_DblClick()

    If List1.ListIndex = -1 Then Exit Sub
    List2.AddItem List1.List(List1.ListIndex)
    List1.RemoveItem List1.ListIndex

End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    nx = (13 - Int(13 - (Picture5.ScaleY(Y, 1, 3) / Picture5.TextHeight("ql"))) - 1)
    nx = nx + List1.TopIndex
    If nx < 0 Then Exit Sub
    If Picture5.TextWidth(List1.List(nx)) > (List1.Width - 22) Then
        List1.ToolTipText = List1.List(nx)
    Else
        List1.ToolTipText = ""
    End If

End Sub


Private Sub List2_DblClick()

    If List2.ListIndex = -1 Then Exit Sub
    List1.AddItem List2.List(List2.ListIndex)
    List2.RemoveItem List2.ListIndex

End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    nx = (13 - Int(13 - (Picture5.ScaleY(Y, 1, 3) / Picture5.TextHeight("ql"))) - 1)
    nx = nx + List2.TopIndex
    If nx < 0 Then Exit Sub
    If Picture5.TextWidth(List2.List(nx)) > (List2.Width - 22) Then
        List2.ToolTipText = List2.List(nx)
    Else
        List2.ToolTipText = ""
    End If

End Sub

Private Sub List3_Click()

    On Error Resume Next
    If ActiveChannel = 9 And Check4 = 1 Then Exit Sub
    dmPerf.SendPatchPMSG 0, DMUS_PMSGF_REFTIME, ActiveChannel, List3.ListIndex, 0, 0
    DoEvents
    ChannelDesc(ActiveChannel).Instrument = List3.ListIndex
    ChannelDesc(ActiveChannel).PatchChanged = True
    Label7 = List3
    

End Sub
Private Sub GetTextLyric()
    
    Text1.Move 0, 0, 649, 223
    Text1 = vbTab & Label10(0) & Chr(13) & Chr(10)
    Text1 = Text1 & Chr(13) & Chr(10)
    For t = 1 To UBound(Phrase)
        Text1 = Text1 & vbTab & Phrase(t) & Chr(13) & Chr(10)
    Next

End Sub
Private Sub Form_Load()
    
    'Center Form
    Left = (((Screen.Width / Screen.TwipsPerPixelX) / 2) - (ScaleX(Width, 1, 3) / 2)) * Screen.TwipsPerPixelX
    Top = (((Screen.Height / Screen.TwipsPerPixelY) / 2) - (ScaleY(Height, 1, 3) / 2)) * Screen.TwipsPerPixelY
    
    'Initialize slider class
    Volume.SlideCreate VolumeCtl
    Balance.SlideCreate BalanceCtl
    Velocity.SlideCreate VelocityCtl
    Transpose.SlideCreate TransposeCtl
    Volume.Value = 64
    Balance.Value = 64
    Velocity.Value = 64
    Transpose.Value = 64
    For t = 0 To 15
        Sld(t).SlideCreate Picture1(t)
    Next
    
    SetPreferences
    
    'Initialize directX
    Set dmLoader = Dx.DirectMusicLoaderCreate

    
    LoadPatches
    Set Picture5.Font = List2.Font
    
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Drag form
    If (X > 12 And X < 721) And (Y > 9 And Y < 28) Then
        If Button = 1 Then
            Dim lngReturnValue As Long
            Call ReleaseCapture
            lngReturnValue = SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
        Exit Sub
    End If
    
    'Progress Bar
    If ChangePosition Then
        If Button = 1 Then
            X = X - X1
            If X < 305 Then X = 305
            If X > 523 - Image3.Width Then X = 523 - Image3.Width
            Image3.Left = X
        End If
        DoEvents
        Exit Sub
    End If
    
    'Preferences button
    If (X > 595 And X < 709) And (Y > 277 And Y < 292) Then
        Label2(2) = "Preferences"
        Exit Sub
    End If
    
    'List Buttons
    If (X > 604 And X < 745) And (Y > 452 And Y < 471) Then
        Select Case Picture6.Point(X - 605, Y - 453)
            Case RGB(255, 0, 0)
                Label2(2) = "Previous Music"
            Case RGB(0, 255, 0)
                Label2(2) = "First Music"
            Case RGB(0, 0, 255)
                Label2(2) = "Last Music"
            Case RGB(255, 0, 255)
                Label2(2) = "Next Music"
            Case Else
                Label2(2) = ""
        End Select
        Exit Sub
    End If
    
    
    'Master buttons
    Select Case Picture2.Point(X - 599, Y - 484)
        Case RGB(255, 0, 0)
            Label2(2) = GetTip(1)
        Case RGB(255, 255, 0)
            Label2(2) = GetTip(2)
        Case RGB(0, 255, 0)
            Label2(2) = GetTip(3)
        Case RGB(0, 0, 255)
            Label2(2) = GetTip(4)
        Case RGB(255, 0, 255)
            Label2(2) = GetTip(5)
        Case RGB(255, 255, 255)
            Label2(2) = GetTip(6)
        Case RGB(0, 0, 0)
            Label2(2) = GetTip(7)
        Case RGB(0, 255, 255)
            Label2(2) = GetTip(8)
        Case Else
            Label2(2) = ""
    End Select
    Label2(1) = ""
    
End Sub
Private Sub Label11_Click()

    On Error Resume Next
    If Not DeviceOpen Then Exit Sub
    If List3.Visible Then
        List3.Visible = False
    Else
        List3.Visible = True
        List3.SetFocus
    End If

End Sub
Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Or KeyCode = 13 Then List3.Visible = False

End Sub
Private Sub List3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then List3.Visible = False

End Sub
Private Sub Option1_Click(Index As Integer)
    
    On Error Resume Next
    If Not DeviceOpen Then Exit Sub
    List2.SetFocus
    Label11.Enabled = True
    ActiveChannel = Index
    Label6 = "Channel " & Format(Index + 1, "00")
    List3 = GetPatch(ChannelDesc(Index).Instrument)
    Option2(ChannelDesc(Index).Octave) = True
    
End Sub
Private Sub Option2_Click(Index As Integer)
    
    On Error Resume Next
    
    If ActiveChannel = 9 And Check4 = 1 Then Exit Sub
    Select Case Index
        Case 0
            Vlr = -24
        Case 1
            Vlr = -12
        Case 2
            Vlr = 0
        Case 3
            Vlr = 12
        Case 4
            Vlr = 24
    End Select
    Vlr = Vlr + (Int((((24 / 127) * Transpose.Value) - 12) / 2) * 2)
    ChannelDesc(ActiveChannel).Octave = Index
    dmPerf.SendTransposePMSG 0, DMUS_PMSGF_REFTIME, ActiveChannel, Vlr
    ChannelDesc(ActiveChannel).PatchChanged = True
    DoEvents

End Sub
Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Sld(Index).DragButton Button, X, Y
    If Button = 1 Then
        dmPerf.SendMIDIPMSG 0, DMUS_PMSGF_REFTIME, Index, &HB0, &H7, Sld(Index).Volume
        dmPerf.SendMIDIPMSG 0, DMUS_PMSGF_REFTIME, Index, &HB0, &HA, Sld(Index).Balance
        ChannelDesc(Index).Volume = Sld(Index).Volume
        ChannelDesc(Index).Balance = Sld(Index).Balance
        ChannelDesc(Index).VolumeChanged = True
    End If
    DoEvents
    
End Sub

Private Sub Picture5_Click()

    If Not LstMouseOver Or LstButtonIndex = 0 Then Exit Sub
    
    Select Case LstButtonIndex
        Case 1
            If List1.ListIndex = -1 Then Exit Sub
            List2.AddItem List1.List(List1.ListIndex)
            List1.RemoveItem List1.ListIndex
        Case 2
            If List1.ListCount = 0 Then Exit Sub
            For t = 0 To List1.ListCount - 1
                List2.AddItem List1.List(t)
            Next
            List1.Clear
        Case 3
            If List2.ListIndex = -1 Then Exit Sub
            List1.AddItem List2.List(List2.ListIndex)
            List2.RemoveItem List2.ListIndex
        Case 4
            If List2.ListCount = 0 Then Exit Sub
            For t = 0 To List2.ListCount - 1
                List1.AddItem List2.List(t)
            Next
            List2.Clear
        Case 5
            OpenList
        Case 6
            SaveList
        Case 7
            BrowseForKar
        Case 8
            List2.Clear
            FileName = ""
        Case 9
            Picture5.Visible = False
            FirstList
            SeekStart
            PanelEnabled = True
    End Select
    LstButtonIndex = 0
    

End Sub
Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        LstButtonIndex = 0
        
        'Move Buttons
        If (X > 260 And X < 301) And (Y > 52 And Y < 150) Then
            SetCapture Picture5.hWnd
            Select Case Picture7.Point(X - 260, Y - 52)
                Case RGB(255, 0, 0)
                    Picture5.PaintPicture Image9, 260, 54, , , 1, 1, 42, 22
                    LstButtonIndex = 1
                Case RGB(0, 255, 0)
                    Picture5.PaintPicture Image9, 260, 76, , , 1, 23, 42, 22
                    LstButtonIndex = 2
                Case RGB(0, 0, 255)
                    Picture5.PaintPicture Image9, 260, 105, , , 1, 52, 42, 22
                    LstButtonIndex = 3
                Case RGB(255, 0, 255)
                    Picture5.PaintPicture Image9, 260, 127, , , 1, 74, 42, 22
                    LstButtonIndex = 4
            End Select
            Exit Sub
        End If
    
        'Save, Open, Browse ... buttons
        If (X > 563 And X < 640) And (Y > 10 And Y < 214) Then
            SetCapture Picture5.hWnd
            Select Case Picture7.Point((X + 45) - 563, Y - 10)
                Case RGB(255, 0, 0)
                    Picture5.PaintPicture Image9, 563, 11, , , 45, 1, 77, 30
                    LstButtonIndex = 5
                Case RGB(0, 255, 0)
                    Picture5.PaintPicture Image9, 563, 42, , , 45, 32, 77, 30
                    LstButtonIndex = 6
                Case RGB(0, 0, 255)
                    Picture5.PaintPicture Image9, 563, 74, , , 45, 64, 77, 30
                    LstButtonIndex = 7
                Case RGB(255, 0, 255)
                    Picture5.PaintPicture Image9, 563, 155, , , 45, 145, 77, 30
                    LstButtonIndex = 8
                Case RGB(0, 255, 255)
                    Picture5.PaintPicture Image9, 563, 187, , , 45, 177, 77, 30
                    LstButtonIndex = 9
            End Select
        End If
    End If

End Sub
Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label2(2) = ""

End Sub
Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If LstButtonIndex = 0 Then Exit Sub
    LstMouseOver = False
    
    If Button = 1 Then
        'Move Buttons
        Select Case LstButtonIndex
            Case 1
                Picture5.PaintPicture Image10, 260, 54, , , 1, 1, 42, 22
                If Picture7.Point(X - 260, Y - 52) = RGB(255, 0, 0) Then LstMouseOver = True
            Case 2
                Picture5.PaintPicture Image10, 260, 76, , , 1, 23, 42, 22
                If Picture7.Point(X - 260, Y - 52) = RGB(0, 255, 0) Then LstMouseOver = True
            Case 3
                Picture5.PaintPicture Image10, 260, 105, , , 1, 52, 42, 22
                If Picture7.Point(X - 260, Y - 52) = RGB(0, 0, 255) Then LstMouseOver = True
            Case 4
                Picture5.PaintPicture Image10, 260, 127, , , 1, 74, 42, 22
                If Picture7.Point(X - 260, Y - 52) = RGB(255, 0, 255) Then LstMouseOver = True
        End Select
        
        'Save, Open, Browse ... buttons
        Select Case LstButtonIndex
            Case 5
                Picture5.PaintPicture Image10, 563, 11, , , 45, 1, 77, 30
                If Picture7.Point((X + 45) - 563, Y - 10) = RGB(255, 0, 0) Then LstMouseOver = True
            Case 6
                Picture5.PaintPicture Image10, 563, 42, , , 45, 32, 77, 30
                If Picture7.Point((X + 45) - 563, Y - 10) = RGB(0, 255, 0) Then LstMouseOver = True
            Case 7
                Picture5.PaintPicture Image10, 563, 74, , , 45, 64, 77, 30
                If Picture7.Point((X + 45) - 563, Y - 10) = RGB(0, 0, 255) Then LstMouseOver = True
            Case 8
                Picture5.PaintPicture Image10, 563, 155, , , 45, 145, 77, 30
                If Picture7.Point((X + 45) - 563, Y - 10) = RGB(255, 0, 255) Then LstMouseOver = True
            Case 9
                Picture5.PaintPicture Image10, 563, 187, , , 45, 177, 77, 30
                If Picture7.Point((X + 45) - 563, Y - 10) = RGB(0, 255, 255) Then LstMouseOver = True
        End Select
    End If
    ReleaseCapture

End Sub

Private Sub Timer1_Timer()

    On Error Resume Next
    Select Case Status
        Case 0
            Label12 = "Playing"
        Case 1
            Label12 = "Stopped"
            Exit Sub
        Case 2
            Label12 = "Paused"
            Exit Sub
    End Select
    
    Dim DeltaTime As Double
    If Xini = 0 Then Xini = 1
    DoEvents
    s = (dmPerf.GetMasterTempo * 1000) - Bpm
    DeltaTime = (((dmState.GetSeek / 768000) * ppqn) - s)
    If Xini + 1 <= UBound(Lyr) Then
        If DeltaTime >= Lyr(Xini).TempoTotal Then
            Xini = Xini + 1
            Tw1 = (Picture4.Width / 2) - (Picture4.TextWidth(Phrase(Lyr(Xini).FraseIndex)) / 2)
            Tw2 = (Picture4.Width / 2) - (Picture4.TextWidth(Phrase(Lyr(Xini).FraseIndex + 1)) / 2)
            If ShowBall Then
                If Trim(Phrase(Lyr(Xini).FraseIndex)) <> "" Then
                    Timer2.Enabled = True
                    Shape1.Visible = True
                Else
                    Timer2.Enabled = False
                    Shape1.Visible = False
                End If
            End If
            
            Picture4.Cls
            Picture4.CurrentY = 68
            Picture4.CurrentX = Tw1
            Picture4.ForeColor = Crl1
            Picture4.Print Left(Phrase(Lyr(Xini).FraseIndex), Lyr(Xini).TextStart);
            
            Picture4.ForeColor = Crl2
            Picture4.Print Lyr(Xini).TxtString;
            
            Picture4.ForeColor = Crl1
            Rst = Len(Left(Phrase(Lyr(Xini).FraseIndex), Lyr(Xini).TextStart)) + Lyr(Xini).TxtStringLen
            Picture4.Print Right(Phrase(Lyr(Xini).FraseIndex), Len(Phrase(Lyr(Xini).FraseIndex)) - Rst)

            Picture4.CurrentY = Picture4.CurrentY + 12
            Picture4.CurrentX = Tw2
            Picture4.Print Phrase(Lyr(Xini).FraseIndex + 1)
            
            Shape1.Left = Tw1 + Picture4.TextWidth(Left(Phrase(Lyr(Xini).FraseIndex), Lyr(Xini).TextStart)) + ((Picture4.TextWidth(Lyr(Xini).TxtString) - Shape1.Width) / 2)
            Shape1.Top = 48
            DoEvents
        End If
    End If

    If Status <> 0 Then Picture3.Cls
    Sec = Int(dmState.GetSeek / 1000)
    Min = Int(Sec / 60)
    Sec = Int(Sec - (Min * 60))
    Label4 = Format(Min, "00") & ":" & Format(Sec, "00")
    If dmState.GetSeek >= dmSeg.GetLength Then
        SeekEnd
    End If
    
    Vle = Int((dmState.GetSeek * 100) / dmSeg.GetLength)
    Label8 = Vle & " %"
    DoEvents
    UpdateBar Vle

    If dmPerf.GetNotificationPMSG(msg) Then
        Label1 = msg.lField1 + 1
    End If
    
    
    


End Sub
Private Sub Timer2_Timer()

    Static Flag As Boolean
    PosY = Shape1.Top
    If Not Flag Then
        PosY = PosY + 4
        If PosY > 48 Then
            PosY = PosY - 8
            Flag = True
        End If
    Else
        PosY = PosY - 4
        If PosY < 24 Then
            PosY = PosY + 8
            Flag = False
        End If
    End If
    Shape1.Top = PosY

End Sub

Private Sub TransposeCtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Transpose.DragButton Button, X, Y
    Label2(2) = "Transpose"
    If Button = 1 Then
        For t = 0 To 15
            If ChannelDesc(t).Channel <> -1 Then
                If t <> 9 Then
                    dmPerf.SendTransposePMSG 0, DMUS_PMSGF_REFTIME, t, Int((((24 / 127) * Transpose.Value) - 12) / 2) * 2
                Else
                    If Check4 = 0 Then
                        dmPerf.SendTransposePMSG 0, DMUS_PMSGF_REFTIME, t, Int((((24 / 127) * Transpose.Value) - 12) / 2) * 2
                    End If
                End If
                DoEvents
            End If
        Next
    End If
    Vlr = (Int((((24 / 127) * Transpose.Value) - 12) / 2) * 2) / 2
    If Vlr <> 0 Then
        Label2(1) = IIf(Vlr < 0, Vlr, "+" & Vlr)
    Else
        Label2(1) = "0"
    End If
    
End Sub
Private Sub VelocityCtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Velocity.DragButton Button, X, Y
    Label2(2) = "Velocity"
    Vlr = Int((30 / 127) * Velocity.Value)
    If Button = 1 Then
        dmPerf.SetMasterTempo ((2 / 30) * Vlr)
    End If
    If Vlr <> 15 Then
        Label2(1) = IIf(Vlr < 15, "-" & 15 - Vlr, "+" & Vlr - 15)
    Else
        Label2(1) = "0"
    End If

End Sub
Private Sub VolumeCtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Volume.DragButton Button, X, Y
    Label2(2) = "Volume"
    If Button = 1 Then
        dmPerf.SetMasterVolume ((Volume.Value * 42) - 3000)
    End If
    Label2(1) = Int((30 / 127) * Volume.Value)

End Sub


