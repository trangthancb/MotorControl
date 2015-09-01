VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{EB7A6012-79A9-4A1A-91AF-F2A92FCA3406}#1.0#0"; "TeeChart8.ocx"
Object = "{1D8AB547-1323-4FDA-BEDB-A2759F814B83}#269.0#0"; "UnicodeFullControl.ocx"
Begin VB.Form Form1 
   ClientHeight    =   9795
   ClientLeft      =   2790
   ClientTop       =   855
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "DCTDDC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   14850
   Begin VB.CommandButton cmd_bat 
      Caption         =   "Bat Led"
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
      Left            =   9720
      TabIndex        =   0
      Top             =   6000
      Width           =   975
   End
   Begin UnicodeControl.UniFrame UniFrame5 
      Height          =   2535
      Left            =   120
      TabIndex        =   25
      Top             =   7200
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   4471
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OldHeight       =   2535
      UniToolTipText  =   "DCTDDC.frx":6852
      Caption         =   "DCTDDC.frx":686A
      ForeColor       =   12583104
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Picture         =   "DCTDDC.frx":689C
      PictureType     =   0
      Begin UnicodeControl.UniLabel UniLabel11 
         Height          =   345
         Left            =   1080
         TabIndex        =   31
         Top             =   2160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":68B4
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":68EA
         Picture         =   "DCTDDC.frx":6902
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel10 
         Height          =   345
         Left            =   1080
         TabIndex        =   30
         Top             =   1800
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   609
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":691A
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":6950
         Picture         =   "DCTDDC.frx":6968
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel9 
         DragIcon        =   "DCTDDC.frx":6980
         Height          =   345
         Left            =   1080
         TabIndex        =   29
         Top             =   1440
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D1D2
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D20A
         Picture         =   "DCTDDC.frx":D222
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel8 
         Height          =   435
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   767
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D23A
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D264
         Picture         =   "DCTDDC.frx":D27C
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel7 
         Height          =   435
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   767
         ForeColor       =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D294
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D2F0
         Picture         =   "DCTDDC.frx":D308
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel6 
         Height          =   435
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   14160
         _ExtentX        =   24977
         _ExtentY        =   767
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D320
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D43C
         Picture         =   "DCTDDC.frx":D454
         PictureType     =   0
      End
   End
   Begin VB.CommandButton cmd_tat 
      Caption         =   "Tat Led"
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
      Left            =   9720
      TabIndex        =   1
      Top             =   6480
      Width           =   975
   End
   Begin VB.Timer tmr_seach_com 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   9360
      Top             =   3360
   End
   Begin UnicodeControl.UniFrame UniFrame4 
      Height          =   2775
      Left            =   11040
      TabIndex        =   17
      Top             =   4320
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OldHeight       =   2775
      UniToolTipText  =   "DCTDDC.frx":D46C
      Caption         =   "DCTDDC.frx":D484
      ForeColor       =   -2147483635
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Picture         =   "DCTDDC.frx":D4BC
      PictureType     =   0
      Begin UnicodeControl.UniButton cmd_cauhinh 
         Height          =   495
         Left            =   1080
         TabIndex        =   24
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "DCTDDC.frx":D4D4
         ForecolorSelected=   0
         UniToolTipText  =   "DCTDDC.frx":D514
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableDoubleClick=   -1  'True
         Picture         =   "DCTDDC.frx":D52C
      End
      Begin UnicodeControl.UniTextBox txt_Kd 
         Height          =   405
         Left            =   2040
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Alignment       =   2
         UniToolTipText  =   "DCTDDC.frx":D544
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   2
         CaptureEnter    =   0   'False
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   -1
         MouseIcon       =   "DCTDDC.frx":D55C
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "DCTDDC.frx":D578
         UseEvents       =   -1  'True
      End
      Begin UnicodeControl.UniTextBox txt_Ki 
         Height          =   405
         Left            =   2040
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Alignment       =   2
         UniToolTipText  =   "DCTDDC.frx":D59A
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   2
         CaptureEnter    =   0   'False
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   -1
         MouseIcon       =   "DCTDDC.frx":D5B2
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "DCTDDC.frx":D5CE
         UseEvents       =   -1  'True
      End
      Begin UnicodeControl.UniTextBox txt_Kp 
         Height          =   405
         Left            =   2040
         TabIndex        =   21
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Alignment       =   2
         UniToolTipText  =   "DCTDDC.frx":D5F8
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   2
         CaptureEnter    =   0   'False
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   -1
         MouseIcon       =   "DCTDDC.frx":D610
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "DCTDDC.frx":D62C
         UseEvents       =   -1  'True
      End
      Begin UnicodeControl.UniLabel UniLabel5 
         Height          =   345
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D64E
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D674
         Picture         =   "DCTDDC.frx":D68C
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel4 
         Height          =   345
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D6A4
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D6CA
         Picture         =   "DCTDDC.frx":D6E2
         PictureType     =   0
      End
      Begin UnicodeControl.UniLabel UniLabel3 
         Height          =   345
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D6FA
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D720
         Picture         =   "DCTDDC.frx":D738
         PictureType     =   0
      End
   End
   Begin UnicodeControl.UniFrame UniFrame3 
      Height          =   1935
      Left            =   11040
      TabIndex        =   9
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OldHeight       =   1935
      UniToolTipText  =   "DCTDDC.frx":D750
      Caption         =   "DCTDDC.frx":D768
      ForeColor       =   -2147483635
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Picture         =   "DCTDDC.frx":D7B2
      PictureType     =   0
      Begin UnicodeControl.UniTextBox txt_Vthuc 
         Height          =   405
         Left            =   1560
         TabIndex        =   16
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         Alignment       =   2
         UniToolTipText  =   "DCTDDC.frx":D7CA
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   2
         CaptureEnter    =   0   'False
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         HideSelection   =   -1  'True
         Locked          =   -1  'True
         MaxLength       =   -1
         MouseIcon       =   "DCTDDC.frx":D7E2
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "DCTDDC.frx":D7FE
         UseEvents       =   -1  'True
      End
      Begin UnicodeControl.UniTextBox txt_Vdat 
         Height          =   405
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   714
         Alignment       =   2
         UniToolTipText  =   "DCTDDC.frx":D826
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   2
         CaptureEnter    =   0   'False
         CaptureEsc      =   0   'False
         CaptureTab      =   0   'False
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   -1
         MouseIcon       =   "DCTDDC.frx":D83E
         MousePointer    =   0
         MultiLine       =   0   'False
         PasswordChar    =   ""
         RightToLeft     =   0   'False
         ScrollBars      =   0
         Text            =   "DCTDDC.frx":D85A
         UseEvents       =   -1  'True
      End
      Begin UnicodeControl.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
      End
      Begin UnicodeControl.UniLabel UniLabel2 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D880
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D8BE
         Picture         =   "DCTDDC.frx":D8D6
         PictureType     =   0
      End
      Begin UnicodeControl.UniButton cmd_setV 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "DCTDDC.frx":D8EE
         ForecolorSelected=   0
         UniToolTipText  =   "DCTDDC.frx":D914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableDoubleClick=   -1  'True
         Picture         =   "DCTDDC.frx":D92C
      End
      Begin UnicodeControl.UniLabel UniLabel1 
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":D944
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":D980
         Picture         =   "DCTDDC.frx":D998
         PictureType     =   0
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   9360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   6
      DTREnable       =   -1  'True
      InputLen        =   1
      RThreshold      =   1
   End
   Begin UnicodeControl.UniFrame UniFrame2 
      Height          =   2415
      Left            =   11040
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OldHeight       =   2415
      UniToolTipText  =   "DCTDDC.frx":D9B0
      Caption         =   "DCTDDC.frx":D9C8
      ForeColor       =   -2147483635
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Picture         =   "DCTDDC.frx":D9FA
      PictureType     =   0
      Begin UnicodeControl.ProgressBar pro_seach_com 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         Max             =   50
      End
      Begin UnicodeControl.UniLabel lbl_thongbao 
         Height          =   345
         Left            =   240
         TabIndex        =   32
         Top             =   1320
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AutoSize        =   -1  'True
         Caption         =   "DCTDDC.frx":DA12
         Link            =   "myspecialbox@yahoo.com.vn"
         ForeColorWordEffect=   0
         SpeedOrtherColor=   0
         UniToolTipText  =   "DCTDDC.frx":DA4E
         Picture         =   "DCTDDC.frx":DA66
         PictureType     =   0
      End
      Begin UnicodeControl.UniCombobox cb_com 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Text            =   "DCTDDC.frx":DA7E
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   0
         UniToolTipText  =   "DCTDDC.frx":DAA4
         Icon            =   -1
         Item            =   "DCTDDC.frx":DABC
         ComboStyle      =   0
      End
      Begin UnicodeControl.UniButton cmd_connect 
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         Caption         =   "DCTDDC.frx":DB16
         ForecolorSelected=   0
         UniToolTipText  =   "DCTDDC.frx":DB48
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableDoubleClick=   -1  'True
         Picture         =   "DCTDDC.frx":DB60
      End
      Begin UnicodeControl.UniOption op_thucong 
         Height          =   345
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   609
         Caption         =   "DCTDDC.frx":DB78
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   0
         AutoSize        =   -1  'True
         UniToolTipText  =   "DCTDDC.frx":DBAA
         Theme           =   ""
      End
      Begin UnicodeControl.UniOption op_tudong 
         Height          =   345
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
         Value           =   -1  'True
         Caption         =   "DCTDDC.frx":DBC2
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   0
         AutoSize        =   -1  'True
         UniToolTipText  =   "DCTDDC.frx":DBF4
         Theme           =   ""
      End
   End
   Begin UnicodeControl.UniFrame UniFrame1 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12938
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OldHeight       =   7335
      UniToolTipText  =   "DCTDDC.frx":DC0C
      Caption         =   "DCTDDC.frx":DC24
      ForeColor       =   -2147483635
      ForeHighlight   =   0
      ShowButton      =   0   'False
      Picture         =   "DCTDDC.frx":DC58
      PictureType     =   0
      Begin TeeChart.TChart TChart1 
         Height          =   6735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   10695
         Base64          =   $"DCTDDC.frx":DC70
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_giatri As String 'Bien trung gian luu gia tri doc tu PIC
Dim seach_COM As Integer 'Bien trung gian de tim thiet bi
Dim COM_true As Boolean 'Bien trung gian xac dinh dung cong COM
Dim dien_ap As Integer
Dim rs_kind As Byte
Dim tocdo(1 To 50000) As Long
Dim temp_tocdo As Long

Private Sub Command1_Click()
    'lbl_udk.Caption = CStr(14 - (HScroll1.Value * 0.0546875))
           
            Dim i As Integer
            For i = 1 To Len(Text2.Text)
                MSComm1.Output = Mid(Text2, i, 1)
            Next i
             MSComm1.Output = "u#"
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub

Private Sub cmd_setV_Click()
    MSComm1.Output = txt_Vdat + "v#"
End Sub

'==================================
'Cai dat khi Load chuong trinh
'==================================
Private Sub Form_Load()
    seach_COM = 0
    COM_true = False
    dien_ap = 0
    rs_kind = 0
    temp_tocdo = 1
    rs_giatri = ""
End Sub

'==================================
'Doan chuong trinh ket noi cong COM
'==================================
Private Sub cmd_connect_Click()
On Error GoTo Loi
If cmd_connect.Caption = "Kêìt Nôìi" Then
    If op_tudong.Value = True Then
        tmr_seach_com.Enabled = True
        pro_seach_com.Value = 0
        pro_seach_com.Visible = True
    Else
        If MSComm1.PortOpen = False Then 'Kiem tra chua ket noi cong COM
            MSComm1.Settings = "9600,n,8,1"
            If cb_com.Text = "COM" Then
                MsgBox "Baòn chýa choòn côÒng COM", , "Thông baìo"
                Exit Sub
            End If
            MSComm1.CommPort = cb_com.Text
            MSComm1.PortOpen = True 'Mo ket noi cong COM
            lbl_thongbao.Caption = "ÐaÞ kêìt nôìi COM" + cb_com.Text
            cb_com.Enabled = False
        Else
            MsgBox "ÐaÞ kêìt nôìi rôÌi", , "Thông baìo"
        End If
        cmd_connect.Caption = "Ngãìt kêìt nôìi"
    End If
Else
    cmd_connect.Caption = "Kêìt Nôìi"
    lbl_thongbao.Caption = "ÐaÞ ngãìt kêìt nôìi"
    cb_com.Enabled = True
    MSComm1.PortOpen = False
    COM_true = False
    seach_COM = 0
End If
Exit Sub
Loi:
    lbl_thongbao.Caption = "Kêìt nôìi lôÞi..."
End Sub

Private Sub HScroll1_Change()
    'MsgBox HScroll1.Value
    lbl_udk.Caption = CStr(12 - (HScroll1.Value * 0.046875))
    If HScroll1.Value < 10 Then
        MSComm1.Output = "00" + CStr(HScroll1.Value) + "u#"
    Else
        If HScroll1.Value < 100 Then
            MSComm1.Output = "0" + CStr(HScroll1.Value) + "u#"
        Else
            MSComm1.Output = CStr(HScroll1.Value) + "u#"
        End If
    End If
    
End Sub

Private Sub Timer1_Timer()
If dien_ap = 225 Then
    dien_ap = 0
    Timer1.Enabled = False
End If
MSComm1.Output = CStr(256 - dien_ap) + "u#"
dien_ap = dien_ap + 1
End Sub

'==================================
'Tu dong do thiet bi
'==================================
Private Sub tmr_seach_com_Timer()
On Error GoTo Handler
    tmr_seach_com.Interval = 400
    If COM_true = True Then
        lbl_thongbao.Caption = "ÐaÞ kêìt nôìi COM" + CStr(seach_COM)
        cb_com.Enabled = False
        cmd_connect.Caption = "Ngãìt kêìt nôìi"
        tmr_seach_com.Enabled = False
        pro_seach_com.Visible = False
        Exit Sub
    Else
        seach_COM = seach_COM + 1
        pro_seach_com.Value = seach_COM
        If seach_COM = 50 Then seach_COM = 1
        lbl_thongbao.Caption = "Ðang thýÒ kêìt nôìi COM" + CStr(seach_COM)
        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    End If
    
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.CommPort = CStr(seach_COM)
    MSComm1.PortOpen = True 'Mo ket noi cong COM
    MSComm1.Output = "&#"
    Exit Sub
Handler:
    tmr_seach_com.Interval = 100
End Sub

'==================================
'Tuy chon Tu dong/Thu cong
'==================================
Private Sub op_tudong_Click()
    If op_tudong.Value = True Then
        cb_com.Visible = False
    End If
End Sub
Private Sub op_thucong_Click()
    If op_thucong.Value = True Then
        cb_com.Visible = True
    End If
End Sub

'==================================
'Cap nhat ty gia
'==================================
Private Sub cmd_capnhat_Click()
    Dim i As Integer
    If MSComm1.PortOpen = False Then
        MsgBox "Chua ket noi voi thiet bi", , "Thong bao"
        Exit Sub
    End If
    If Len(txt_capnhat) > 4 Then txt_capnhat = Left(txt_capnhat, 5)
    MSComm1.Output = "#"
    For i = 1 To 5
        MSComm1.Output = Mid(txt_capnhat, i, 1)
    Next i
    MSComm1.Output = "@"
    txt_capnhat.Visible = False
    cmd_capnhat.Visible = False
    lbl_tygia.Visible = True
    lbl_tygia.Caption = txt_capnhat
End Sub

'==================================
'Bat tat led don
'==================================
Private Sub cmd_bat_Click()
    MSComm1.Output = ")#"
End Sub
Private Sub cmd_tat_Click()
    MSComm1.Output = "(#"
End Sub

'==================================
'Click sua chua ty gia
'==================================
Private Sub lbl_tygia_DblClick()
    txt_capnhat.Visible = True
    cmd_capnhat.Visible = True
    lbl_tygia.Visible = False
    txt_capnhat.SetFocus
    txt_capnhat = lbl_tygia.Caption
End Sub
'==================================
'Ngat ky tu nhan duoc tu PIC
'==================================
Private Sub MSComm1_OnComm()
    Dim kytunhan As String 'Khai bao bien
    kytunhan = MSComm1.Input 'Doc gia tri nhan duoc tu cong com
    If kytunhan = "&" Then
        COM_true = True
        Exit Sub
    End If
    If kytunhan = "!" Then
        rs_kind = 1
        Exit Sub
    End If
    If kytunhan = "#" Then
        rs_kind = 2
        Exit Sub
    End If
    If rs_kind <> 0 Then
        If kytunhan = "@" Then
            hienthi_dothi
            Exit Sub
        End If
        rs_giatri = rs_giatri + kytunhan
    End If
        
    'If kytunhan = "@" Then
        'Text1.Text = tygia
        'MANG(bien_dem) = CLng(Val(tygia))
        'If MANG(bien_dem) <> 0 Then
            'With TChart1.Series(0)
             '   .AddXY bien_dem, MANG(bien_dem), "", vbRed
            'End With
           ' bien_dem = bien_dem + 1
       ' End If
       ' tygia = ""
       ' Exit Sub
   ' End If
    'tygia = tygia + kytunhan
    'Me.Caption = kytunhan
End Sub
Private Sub hienthi_dothi()
    Select Case (rs_kind)
    Case 1
        txt_Vthuc = rs_giatri
        tocdo(temp_tocdo) = CLng(Val(rs_giatri))
        If tocdo(temp_tocdo) <> 0 Then
            With TChart1.Series(0)
                .AddXY temp_tocdo, tocdo(temp_tocdo), "", vbRed
            End With
            temp_tocdo = temp_tocdo + 1
        End If
        rs_kind = 0
        rs_giatri = ""
    Case 2
        rs_kind = 0
    End Select
        
End Sub
'==================================
'Update/Save Ty gia
'==================================
Private Sub cmd_update_Click()
    Dim TextLine
    Open "TEST" For Input As #1   ' Mo file.
    Line Input #1, TextLine   ' Doc tung dong gan vao bien TextLine
    lbl_tygia.Caption = Replace(TextLine, """", "")
    txt_capnhat = Replace(TextLine, """", "")
    cmd_capnhat_Click
    Close #1   ' Dong file.
End Sub
Private Sub cmd_save_Click()
    Dim FileNumber
    FileNumber = FreeFile   ' Gan FileNumber = Trong
    Open "TEST" For Output As #FileNumber   ' mo file
    Write #FileNumber, lbl_tygia.Caption  ' Ghi vào file
    Close #FileNumber   ' Dong file
    MsgBox "Da luu du lieu", , "Thong bao"
End Sub
