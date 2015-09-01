VERSION 5.00
Object = "{EB7A6012-79A9-4A1A-91AF-F2A92FCA3406}#1.0#0"; "TeeChart8Eval.ocx"
Begin VB.Form Form1 
   Caption         =   $"p.frx":0000
   ClientHeight    =   9825
   ClientLeft      =   2895
   ClientTop       =   960
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
   Icon            =   "p.frx":0087
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   14850
   Begin VB.Frame Frame5 
      Caption         =   "THÔNG TIN"
      ForeColor       =   &H8000000D&
      Height          =   2430
      Left            =   105
      TabIndex        =   28
      Top             =   7350
      Width           =   10935
      Begin VB.Label Label5 
         Caption         =   "Traàn Thaêng"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   1260
         TabIndex        =   35
         Top             =   1995
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Voõ Vaên Laâm"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1260
         TabIndex        =   34
         Top             =   1680
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "Leâ Vaên Duaån"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1260
         TabIndex        =   33
         Top             =   1365
         Width           =   1680
      End
      Begin VB.Label Label5 
         Caption         =   "doøng ñieän vaø toác ñoä."
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   12
         Left            =   1260
         TabIndex        =   32
         Top             =   630
         Width           =   2730
      End
      Begin VB.Label Label5 
         Caption         =   "SVTH :"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   11
         Left            =   210
         TabIndex        =   31
         Top             =   1575
         Width           =   945
      End
      Begin VB.Label Label5 
         Caption         =   "GVHD: PGS.TS Ñoaøn Quang Vinh"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   10
         Left            =   210
         TabIndex        =   30
         Top             =   945
         Width           =   9975
      End
      Begin VB.Label Label5 
         Caption         =   "ÑEÀ TAØI: Ñieàu chænh toác ñoä ñoäng cô ñieän moät chieàu kích töø ñoäc laäp coù hai khaâu phaûn hoài"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   9
         Left            =   210
         TabIndex        =   29
         Top             =   315
         Width           =   10500
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ÐIÊÌU KHIÊÒN"
      ForeColor       =   &H8000000D&
      Height          =   2430
      Left            =   11130
      TabIndex        =   25
      Top             =   7350
      Width           =   3585
      Begin VB.CommandButton cmd_xoa 
         Caption         =   "Xoìa ðôÌ thiò"
         Height          =   540
         Left            =   945
         TabIndex        =   27
         Top             =   1470
         Width           =   1695
      End
      Begin VB.CommandButton cmd_khoidong 
         Caption         =   "KhõÒi ðôòng"
         Height          =   645
         Left            =   945
         TabIndex        =   26
         Top             =   525
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "THÔNG SÔì"
      ForeColor       =   &H8000000D&
      Height          =   2535
      Left            =   11130
      TabIndex        =   13
      Top             =   4725
      Width           =   3585
      Begin VB.CommandButton cmd_cauhinh 
         Caption         =   "Xuâìt câìu hiÌnh"
         Height          =   435
         Left            =   840
         TabIndex        =   22
         Top             =   1785
         Width           =   1800
      End
      Begin VB.TextBox txt_Ki 
         Alignment       =   2  'Center
         Height          =   405
         Index           =   1
         Left            =   2310
         TabIndex        =   21
         Text            =   "0.001"
         Top             =   1155
         Width           =   855
      End
      Begin VB.TextBox txt_Ki 
         Alignment       =   2  'Center
         Height          =   405
         Index           =   0
         Left            =   1050
         TabIndex        =   20
         Text            =   "0.001"
         Top             =   1155
         Width           =   960
      End
      Begin VB.TextBox txt_Kp 
         Alignment       =   2  'Center
         Height          =   435
         Index           =   1
         Left            =   2310
         TabIndex        =   19
         Text            =   "0.1"
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox txt_Kp 
         Alignment       =   2  'Center
         Height          =   405
         Index           =   0
         Left            =   1050
         TabIndex        =   18
         Text            =   "0.1"
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label5 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2730
         TabIndex        =   17
         Top             =   210
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   1470
         TabIndex        =   16
         Top             =   210
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Ki:"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   105
         TabIndex        =   15
         Top             =   1155
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Kp:"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   105
         TabIndex        =   14
         Top             =   630
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TÔìC ÐÔò (voÌng/phuìt)"
      ForeColor       =   &H8000000D&
      Height          =   2400
      Left            =   11130
      TabIndex        =   5
      Top             =   2310
      Width           =   3615
      Begin VB.TextBox txt_dong 
         Alignment       =   2  'Center
         Height          =   420
         Left            =   1785
         TabIndex        =   12
         Text            =   "0"
         Top             =   1620
         Width           =   900
      End
      Begin VB.CommandButton cmd_setV 
         Caption         =   "Set"
         Height          =   375
         Left            =   2730
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
      Begin VB.TextBox txt_Vthuc 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   1785
         TabIndex        =   9
         Text            =   "0000"
         Top             =   945
         Width           =   885
      End
      Begin VB.TextBox txt_Vdat 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   1785
         TabIndex        =   7
         Text            =   "2000"
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2940
         TabIndex        =   24
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "V/p"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2940
         TabIndex        =   23
         Top             =   945
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Doøng ñieän"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   105
         TabIndex        =   11
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Toác ñoä thöïc:"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Toác ñoä ñaët:"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   105
         TabIndex        =   6
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "KÊìT NÔìI"
      ForeColor       =   &H8000000D&
      Height          =   2310
      Left            =   11130
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmd_connect 
         Caption         =   "Kêìt nôìi"
         Height          =   420
         Left            =   855
         TabIndex        =   4
         Top             =   1755
         Width           =   1995
      End
      Begin VB.ComboBox cb_com 
         Height          =   405
         ItemData        =   "p.frx":12C9
         Left            =   1470
         List            =   "p.frx":12EB
         TabIndex        =   3
         Text            =   "COM"
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.OptionButton op_thucong 
         Caption         =   "ThuÒ công"
         Height          =   285
         Left            =   1890
         TabIndex        =   2
         Top             =   360
         Width           =   1365
      End
      Begin VB.OptionButton op_tudong 
         Caption         =   "Týò ðôòng"
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.Label lbl_thongbao 
         Alignment       =   2  'Center
         Caption         =   "Chöa keát noái..."
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   52
         Top             =   1365
         Width           =   3375
      End
   End
   Begin VB.Timer tmr_seach_com 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   10080
      Top             =   4620
   End
   Begin VB.Frame Form1 
      Caption         =   "ÐÔÌ THIò"
      ForeColor       =   &H8000000D&
      Height          =   7260
      Left            =   105
      TabIndex        =   38
      Top             =   0
      Width           =   10935
      Begin TeeChart.TChart TChart1 
         Height          =   6825
         Left            =   210
         TabIndex        =   39
         Top             =   315
         Width           =   10530
         Base64          =   $"p.frx":130E
         Begin VB.Label Label2 
            Caption         =   "0.4"
            Height          =   375
            Index           =   19
            Left            =   9090
            TabIndex        =   51
            Top             =   5445
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "0.8"
            Height          =   375
            Index           =   18
            Left            =   9090
            TabIndex        =   50
            Top             =   4995
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "2.4"
            Height          =   375
            Index           =   17
            Left            =   9090
            TabIndex        =   49
            Top             =   3285
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "2.0"
            Height          =   375
            Index           =   16
            Left            =   9090
            TabIndex        =   48
            Top             =   3735
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "1.6"
            Height          =   375
            Index           =   15
            Left            =   9090
            TabIndex        =   47
            Top             =   4140
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "1.2"
            Height          =   375
            Index           =   14
            Left            =   9090
            TabIndex        =   46
            Top             =   4590
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "3.6"
            Height          =   375
            Index           =   13
            Left            =   9090
            TabIndex        =   45
            Top             =   1980
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "3.2"
            Height          =   375
            Index           =   11
            Left            =   9090
            TabIndex        =   44
            Top             =   2385
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "2.8"
            Height          =   375
            Index           =   10
            Left            =   9090
            TabIndex        =   43
            Top             =   2835
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "4.0"
            Height          =   375
            Index           =   6
            Left            =   9090
            TabIndex        =   42
            Top             =   1575
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "V/p"
            Height          =   285
            Left            =   495
            TabIndex        =   41
            Top             =   315
            Width           =   510
         End
         Begin VB.Label Label4 
            Caption         =   "A"
            Height          =   285
            Index           =   1
            Left            =   9090
            TabIndex        =   40
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   "V/p"
      Height          =   285
      Left            =   210
      TabIndex        =   37
      Top             =   210
      Width           =   510
   End
   Begin VB.Label Label4 
      Caption         =   "A"
      Height          =   285
      Index           =   0
      Left            =   105
      TabIndex        =   36
      Top             =   105
      Width           =   375
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
Dim vdat As String
Dim temp_rf_khoidong As Boolean

'==================================
'Xuat Kp, Ki
'==================================
Private Sub cmd_cauhinh_Click()
    MSComm1.Output = txt_Kp(0) + "p#" 'Kp toc do
    MSComm1.Output = txt_Ki(0) + "i#" 'Ki toc do
    MSComm1.Output = txt_Kp(1) + "q#" 'Kp dong dien
    MSComm1.Output = txt_Ki(1) + "j#" 'Ki dong dien
End Sub
'==================================
'Khong dong DC
'==================================
Private Sub cmd_khoidong_Click()
    If cmd_khoidong.Caption = "KhõÒi ðôòng" Then
        cmd_khoidong.Caption = "DýÌng ðôòng cõ"
        MSComm1.Output = "0.03d#"
    Else
        cmd_khoidong.Caption = "KhõÒi ðôòng"
        MSComm1.Output = "0d#"
    End If
End Sub
'==================================
'Cai dat toc do
'==================================
Private Sub cmd_setV_Click()
    MSComm1.Output = txt_Vdat + "v#"
    vdat = txt_Vdat
End Sub
'==================================
'Xoa do thi
'==================================
Private Sub cmd_xoa_Click()
    temp_tocdo = 1
    TChart1.Series(0).Clear
    TChart1.Series(1).Clear
    TChart1.Series(2).Clear
    TChart1.Series(0).AddXY 0, 0, "", vbRed
    TChart1.Series(1).AddXY 0, 0, "", vbGreen
    TChart1.Series(2).AddXY 0, 0, "", vbBlue
End Sub
'==================================
'Cai dat khi Load chuong trinh
'==================================
Private Sub Form_Load()
    temp_rf_khoidong = False
    seach_COM = 0
    COM_true = False
    dien_ap = 20
    rs_kind = 0
    temp_tocdo = 1
    rs_giatri = ""
    vdat = "2000"
    TChart1.Series(0).AddXY 0, 0, "", vbRed
    TChart1.Series(1).AddXY 0, 0, "", vbGreen
    TChart1.Series(2).AddXY 0, 0, "", vbBlue
End Sub
'==================================
'Ket noi cong COM
'==================================
Private Sub cmd_connect_Click()
'On Error GoTo Loi
If cmd_connect.Caption = "Kêìt nôìi" Then
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
            lbl_thongbao.Caption = "Ñaõ keát noái COM" + cb_com.Text
            cb_com.Enabled = False
        Else
            MsgBox "ÐaÞ kêìt nôìi rôÌi", , "Thông baìo"
        End If
        cmd_connect.Caption = "Ngãìt kêìt nôìi"
    End If
Else
    cmd_connect.Caption = "Kêìt Nôìi"
    lbl_thongbao.Caption = "Ñaõ ngaét keát noái"
    cb_com.Enabled = True
    MSComm1.PortOpen = False
    COM_true = False
    seach_COM = 0
End If
Exit Sub
Loi:
    lbl_thongbao.Caption = "Keát noái loãi..."
End Sub





'==================================
'Tu dong do cong COM
'==================================
Private Sub tmr_seach_com_Timer()
On Error GoTo Handler
    tmr_seach_com.Interval = 400
    If COM_true = True Then
        lbl_thongbao.Caption = "Ñaõ keát noái COM" + CStr(seach_COM)
        cb_com.Enabled = False
        cmd_connect.Caption = "Ngãìt kêìt nôìi"
        tmr_seach_com.Enabled = False
        pro_seach_com.Visible = False
        Exit Sub
    Else
        seach_COM = seach_COM + 1
        pro_seach_com.Value = seach_COM
        If seach_COM = 50 Then seach_COM = 1
        lbl_thongbao.Caption = "Ñang thöû keát noái COM" + CStr(seach_COM)
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
        pro_seach_com.Visible = True
    End If
End Sub
Private Sub op_thucong_Click()
    If op_thucong.Value = True Then
        cb_com.Visible = True
        pro_seach_com.Visible = False
    End If
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
End Sub
'==================================
'Ve do thi
'==================================
Private Sub hienthi_dothi()
    Select Case (rs_kind)
    Case 1
    Me.Caption = CStr(rs_giatri)
        Dim s() As String
        s = Split(rs_giatri, "|")
        txt_Vthuc = s(0)
        txt_dong = s(1)
        If CStr(Val(s(0))) <> 0 Then
            TChart1.Series(0).AddXY temp_tocdo, CStr(Val(s(0))), "", vbRed
            TChart1.Series(1).AddXY temp_tocdo, CLng(vdat), "", vbGreen
            TChart1.Series(2).AddXY temp_tocdo, CStr(Val(s(1)) * 500), "", vbBlue
            temp_tocdo = temp_tocdo + 1
        End If
        rs_kind = 0
        rs_giatri = ""
    Case 2
        Dim s() As String
        s = Split(rs_giatri, "|")
        If s(1) = "1" Then
            temp_rf_khoidong = True
        End If
        If s(1) = "0" And temp_rf_khoidong = True Then
            temp_rf_khoidong = False
            cmd_khoidong_Click
        End If
        rs_kind = 0
    End Select
End Sub
