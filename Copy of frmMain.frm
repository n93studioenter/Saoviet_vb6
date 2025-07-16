VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   ClientHeight    =   11220
   ClientLeft      =   4035
   ClientTop       =   1635
   ClientWidth     =   17745
   FillColor       =   &H00FD8866&
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Sao Viet Accounting Software"
   ScaleHeight     =   11220
   ScaleWidth      =   17745
   Tag             =   "11"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Height          =   630
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "TaiKhoan"
            Object.ToolTipText     =   "T�i kho�n"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "NgoaiTe"
            Object.ToolTipText     =   "Nguy�n t�"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Kho"
            Object.ToolTipText     =   "Kho"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "VatTu"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "LuuChuyen"
            Object.ToolTipText     =   "L�u chuy�n"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "DuPhong"
            Object.ToolTipText     =   "D� ph�ng"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "TaiSan"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "CN"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   "TongHop"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ThanhPham"
            Object.ToolTipText     =   "Th�nh ph�m"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "KetThuc"
            Object.ToolTipText     =   "Tho�t"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808000&
      Caption         =   "Nh�p ch�ng t�"
      DragIcon        =   "frmMain.frx":57E2
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1440
      MaskColor       =   &H00808000&
      Picture         =   "frmMain.frx":164DC
      Style           =   1  'Graphical
      TabIndex        =   60
      Tag             =   "Voucher"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808000&
      Caption         =   "S� k� to�n"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1440
      MaskColor       =   &H00808000&
      Picture         =   "frmMain.frx":1BCBE
      Style           =   1  'Graphical
      TabIndex        =   59
      Tag             =   "Detail Report"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00808000&
      Caption         =   "B�o c�o thu�& t�i ch�nh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   1440
      MaskColor       =   &H00808000&
      Picture         =   "frmMain.frx":214A0
      Style           =   1  'Graphical
      TabIndex        =   58
      Tag             =   "Financial Report"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   30
      Top             =   4680
      Width           =   3315
      Begin VB.CheckBox chk 
         BackColor       =   &H00808000&
         Caption         =   "TMDV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   0
         MaskColor       =   &H00000000&
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00808000&
         Caption         =   "XD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00808000&
         Caption         =   "CDT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chk 
         BackColor       =   &H00808000&
         Caption         =   "SX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2040
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command 
      Caption         =   "K� to�n      Ch� ��u t�"
      Height          =   210
      Index           =   6
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   500
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command 
      Caption         =   "B�o c�o   Qu�n tr�"
      Height          =   210
      Index           =   5
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   500
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command 
      Height          =   210
      Index           =   4
      Left            =   0
      Picture         =   "frmMain.frx":26C82
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   500
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton Command 
      Caption         =   "Ch��ng tr�nh theo &y�u c�u doanh nghi�p"
      Height          =   210
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "Customized Report"
      Top             =   500
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   5520
      Width           =   2520
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00808000&
         Caption         =   "Vi�t"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   10
         Tag             =   "Vietnamese"
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptNN 
         BackColor       =   &H00808000&
         Caption         =   "Anh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   3
         Left            =   1320
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Tag             =   "English"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.FileListBox File1 
      Height          =   510
      Left            =   8760
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer CTTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   2400
   End
   Begin VB.PictureBox imlIcons 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   11835
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   14
      Top             =   9345
      Visible         =   0   'False
      Width           =   1200
   End
   Begin Crystal.CrystalReport Rpt 
      Left            =   2715
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   0
      Left            =   0
      TabIndex        =   57
      Top             =   11220
      Width           =   17745
      _ExtentX        =   31300
      _ExtentY        =   0
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "05/01/14"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image img 
      Height          =   495
      Left            =   8280
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   61
      Tag             =   "Employee"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Pha�n me�m ke� toa�n Sao Vie�t"
      BeginProperty Font 
         Name            =   "VNI-Lithos"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   690
      Index           =   1
      Left            =   4920
      TabIndex        =   56
      Tag             =   "Accounting Software Company"
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "N�m t�i ch�nh:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   10
      Left            =   7680
      TabIndex        =   55
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   9
      Left            =   7680
      TabIndex        =   54
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "T�nh th�nh:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   8
      Left            =   7680
      TabIndex        =   53
      Tag             =   "Province"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "L�nh v�c:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   7
      Left            =   4440
      TabIndex        =   52
      Tag             =   "Activies"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "M� s� thu�:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   6
      Left            =   4440
      TabIndex        =   51
      Tag             =   "Tax Code"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "S� �i�n tho�i:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   5
      Left            =   4440
      TabIndex        =   50
      Tag             =   "Tel"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "Qu�n huy�n:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   4
      Left            =   4440
      TabIndex        =   49
      Tag             =   "District"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "��a ch�:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   4440
      TabIndex        =   48
      Tag             =   "Address"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "C�ng ty:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   2
      Left            =   4440
      TabIndex        =   47
      Tag             =   "Employee"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "T�n C�ng ty"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   0
      Left            =   5880
      TabIndex        =   46
      Top             =   2400
      Width           =   7095
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "��a ch�"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   2
      Left            =   5880
      TabIndex        =   45
      Top             =   2880
      Width           =   6855
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "T�n C�ng ty"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   10
      Left            =   5880
      TabIndex        =   44
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "�i�n tho�i"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   3
      Left            =   6000
      TabIndex        =   43
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "M� s� thu�"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   8
      Left            =   5880
      TabIndex        =   42
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "T�nh th�nh"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   11
      Left            =   9000
      TabIndex        =   41
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   4
      Left            =   8760
      TabIndex        =   40
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lbCty 
      BackColor       =   &H00808000&
      Caption         =   "N�m t�i ch�nh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Index           =   7
      Left            =   9360
      TabIndex        =   39
      Tag             =   "Financial Year"
      Top             =   4320
      Width           =   2415
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   1320
      TabIndex        =   38
      Top             =   1080
      Width           =   1095
      BackColor       =   8421376
      Size            =   "1931;450"
      BorderColor     =   0
      FontName        =   "Arial"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "��n v� c�i ��t: C�ng ty TNHH DV Thu� Sao Vi�t"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   12
      Left            =   9840
      TabIndex        =   37
      Top             =   7200
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "��a ch�: 640 Tr��ng C�ng ��nh, Tp V�ng T�u"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   13
      Left            =   9840
      TabIndex        =   36
      Top             =   7560
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "S�T: 064.3839 678 - 090 3839 678"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   14
      Left            =   9840
      TabIndex        =   35
      Top             =   7920
      Width           =   3975
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   12
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":27704
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":28A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":29D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2B03A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2C34C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2D65E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2E2B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2FE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":30AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":31EB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":34228
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3461A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      Caption         =   "��n v�"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Index           =   11
      Left            =   9240
      TabIndex        =   29
      Top             =   1800
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 9"
      Height          =   375
      Index           =   9
      Left            =   7920
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label 
      Caption         =   "��n v� ph�t h�nh:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   11400
      TabIndex        =   27
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      Height          =   1215
      Index           =   16
      Left            =   12720
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lbCty 
      AutoSize        =   -1  'True
      Caption         =   "LbCty 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   10200
      TabIndex        =   13
      Tag             =   "0"
      Top             =   2520
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label 
      BackColor       =   &H00808000&
      Caption         =   "Ch��ng tr�nh ����c ca�i �a�t ta�i:"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Tag             =   "This product � is licensed to"
      Top             =   1680
      Width           =   4200
   End
   Begin VB.Label email 
      BackColor       =   &H00808000&
      Height          =   360
      Index           =   10
      Left            =   1800
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   12
      Left            =   11280
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   9240
      TabIndex        =   25
      Top             =   12960
      Width           =   195
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   9240
      TabIndex        =   20
      Top             =   12240
      Width           =   135
   End
   Begin VB.Label Lb 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Index           =   0
      Left            =   9285
      TabIndex        =   19
      Tag             =   "Model"
      Top             =   12105
      Width           =   225
   End
   Begin VB.Label Label 
      Caption         =   "Ng�nh ngh�"
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
      Index           =   25
      Left            =   7920
      TabIndex        =   24
      Tag             =   "Profession"
      Top             =   12960
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Lo�i h�nh DN"
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
      Index           =   23
      Left            =   7920
      TabIndex        =   18
      Tag             =   "Class"
      Top             =   12240
      Width           =   1215
   End
   Begin VB.Label Label 
      Caption         =   "Phi�n b�n"
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
      Index           =   21
      Left            =   7635
      TabIndex        =   17
      Tag             =   "Version"
      Top             =   12225
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   22
      Left            =   11040
      TabIndex        =   5
      Tag             =   "UNESCO Center for Development of Information Technology"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label 
      Caption         =   "T�nh, th�nh ph�"
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
      Left            =   2.45745e5
      TabIndex        =   12
      Tag             =   "Province"
      Top             =   5700
      Width           =   1455
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   22
      Tag             =   "1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lbCty 
      Caption         =   "LbCty 5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   11280
      TabIndex        =   21
      Tag             =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label 
      Caption         =   "T�i kho�n Ngo�i t�"
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
      Index           =   8
      Left            =   2550
      TabIndex        =   23
      Tag             =   "Bank F.C. Account"
      Top             =   12615
      Width           =   1935
   End
   Begin VB.Label Label 
      Caption         =   "T�i kho�n ti�n Vi�t"
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
      Left            =   2535
      TabIndex        =   16
      Tag             =   "Bank VND Account"
      Top             =   12255
      Width           =   1935
   End
   Begin VB.Menu mnuHethong 
      Caption         =   "&H� th�ng"
      Tag             =   "&System"
      WindowList      =   -1  'True
      Begin VB.Menu mnHT 
         Caption         =   "&M� t�p d� li�u..."
         Index           =   0
         Tag             =   "Open data file..."
      End
      Begin VB.Menu mnHT 
         Caption         =   "&Sao ch�p t�p d� li�u..."
         Index           =   1
         Tag             =   "Save as..."
      End
      Begin VB.Menu mnHT 
         Caption         =   "&T�p d� li�u m�c ��nh..."
         Index           =   2
         Tag             =   "Default data file"
      End
      Begin VB.Menu mnHT 
         Caption         =   "&N�n t�p d� li�u..."
         Index           =   3
         Tag             =   "Compress data file..."
      End
      Begin VB.Menu mnHT 
         Caption         =   "M� t�&p d� li�u n�n..."
         Index           =   4
         Tag             =   "Open compressed data file"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnHT 
         Caption         =   "��t m�y &in"
         Index           =   8
         Shortcut        =   ^P
         Tag             =   "Print Setup"
      End
      Begin VB.Menu mnHT 
         Caption         =   "F"
         Index           =   10
         Tag             =   "Font convert"
         Visible         =   0   'False
      End
      Begin VB.Menu mnHT 
         Caption         =   "Th�ng s� h� th�ng"
         Index           =   11
         Tag             =   "Options"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnHT 
         Caption         =   "&Danh s�ch ng��i s� d�ng"
         Index           =   13
         Tag             =   "User List"
      End
      Begin VB.Menu mnHT 
         Caption         =   "��t m�t &kh�u"
         Index           =   14
         Tag             =   "Change Password"
      End
      Begin VB.Menu mnHT 
         Caption         =   "-"
         Index           =   23
      End
      Begin VB.Menu mnHT 
         Caption         =   "��i ng��i s� d�n&g"
         Index           =   24
         Tag             =   "Log off"
      End
      Begin VB.Menu mnHT 
         Caption         =   "K�t th�c c&h��ng tr�nh"
         Index           =   25
         Tag             =   "Quit"
      End
   End
   Begin VB.Menu mnDuLieu 
      Caption         =   "Nh�p s� d� ��u k�"
      Tag             =   "&Tools"
      Begin VB.Menu mnDL 
         Caption         =   "Ki�m tra &nh�p xu�t t�n"
         Index           =   0
         Tag             =   "Inventory Check-Up"
      End
      Begin VB.Menu mnDL 
         Caption         =   "Ki�m tra h� th�ng &t�i kho�n"
         Index           =   1
         Tag             =   "Account Check-Up"
      End
      Begin VB.Menu mnDL 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnDL 
         Caption         =   "X� &l� s� li�u..."
         Index           =   3
         Tag             =   "Run SQL Query..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "&Xo� ph�t sinh th�ng"
         Index           =   4
         Tag             =   "Delete data in month"
         Begin VB.Menu mnXoa 
            Caption         =   "S� d� ��u n�m"
            Index           =   0
         End
         Begin VB.Menu mnXoa 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnXoa 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnXoa 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnXoa 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnXoa 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu mnXoa 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu mnXoa 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu mnXoa 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu mnXoa 
            Caption         =   "9"
            Index           =   9
         End
         Begin VB.Menu mnXoa 
            Caption         =   "10"
            Index           =   10
         End
         Begin VB.Menu mnXoa 
            Caption         =   "11"
            Index           =   11
         End
         Begin VB.Menu mnXoa 
            Caption         =   "12"
            Index           =   12
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnDL 
         Caption         =   "&Chuy�n sang n�m m�i"
         Index           =   6
         Tag             =   "Convert to new Financial Year"
      End
      Begin VB.Menu mnDL 
         Caption         =   "N�&m t�i ch�nh"
         Index           =   7
         Tag             =   "Select Financial Year"
         Begin VB.Menu mnNam 
            Caption         =   "0"
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "1"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "8"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnNam 
            Caption         =   "9"
            Index           =   9
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "Nguy�n t�..."
         Index           =   20
         Tag             =   "Posting Vouchers..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "Ch�ng t� &k�t chuy�n..."
         Index           =   9
         Tag             =   "Posting Vouchers..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "&Ph�n b� chi ph�..."
         Index           =   10
         Tag             =   "Expenses Allocation..."
      End
      Begin VB.Menu mnDL 
         Caption         =   "K�t c&huy�n s� li�u..."
         Index           =   11
         Tag             =   "Monthly Conversion"
      End
      Begin VB.Menu mnDL 
         Caption         =   "Kh&o� s� li�u th�ng"
         Index           =   12
         Tag             =   "Clost data in month"
         Begin VB.Menu mnk 
            Caption         =   "S� d� ��u n�m"
            Index           =   0
         End
         Begin VB.Menu mnk 
            Caption         =   "1"
            Index           =   1
         End
         Begin VB.Menu mnk 
            Caption         =   "2"
            Index           =   2
         End
         Begin VB.Menu mnk 
            Caption         =   "3"
            Index           =   3
         End
         Begin VB.Menu mnk 
            Caption         =   "4"
            Index           =   4
         End
         Begin VB.Menu mnk 
            Caption         =   "5"
            Index           =   5
         End
         Begin VB.Menu mnk 
            Caption         =   "6"
            Index           =   6
         End
         Begin VB.Menu mnk 
            Caption         =   "7"
            Index           =   7
         End
         Begin VB.Menu mnk 
            Caption         =   "8"
            Index           =   8
         End
         Begin VB.Menu mnk 
            Caption         =   "9"
            Index           =   9
         End
         Begin VB.Menu mnk 
            Caption         =   "10"
            Index           =   10
         End
         Begin VB.Menu mnk 
            Caption         =   "11"
            Index           =   11
         End
         Begin VB.Menu mnk 
            Caption         =   "12"
            Index           =   12
         End
      End
      Begin VB.Menu mnDL 
         Caption         =   "Khai b�o m�u bi�u song ng�"
         Index           =   19
         Tag             =   "Financial Report Description"
      End
   End
   Begin VB.Menu mnVatTu 
      Caption         =   "&V�t t�, h�ng ho�"
      Tag             =   "&Product and Contruction Cost"
      Begin VB.Menu mnVT 
         Caption         =   "&Ph�n lo�i v�t t�..."
         Index           =   0
         Tag             =   "Classification..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "�i�m danh v�t t� h�ng ho�..."
         Index           =   17
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "&K�nh ph�n ph�i..."
         Index           =   1
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "L� chuy�n n� b�..."
         Index           =   18
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "Th�nh ph�m ho�n th�nh trong k�..."
         Index           =   19
         Tag             =   "Import-Export Source List..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnVT 
         Caption         =   "Kho..."
         Index           =   20
         Tag             =   "Opeining Balance"
      End
      Begin VB.Menu mnVT 
         Caption         =   "&T�n kho ��u k�..."
         Index           =   3
         Tag             =   "Opeining Balance"
      End
      Begin VB.Menu mnVT 
         Caption         =   "T�nh l�i gi� xu�t kho trong th�ng..."
         Index           =   4
         Tag             =   "Recalculate cost of material in month..."
      End
      Begin VB.Menu mnVT 
         Caption         =   "T�nh gi� v�n h�ng &b�n"
         Index           =   5
         Tag             =   "Recalculate cost of sold gooods"
      End
      Begin VB.Menu mnVT 
         Caption         =   "Ki�m k� t�n kho cu�i &ng�y"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnVT 
         Caption         =   "Ph�n &lo�i c�ng tr�nh, s�n ph�m"
         Index           =   9
         Tag             =   "Classification of Product and Contructions"
      End
      Begin VB.Menu mnVT 
         Caption         =   "&Chi ti�t c�ng tr�nh, s�n ph�m"
         Index           =   10
         Tag             =   "Product and Contruction List"
      End
      Begin VB.Menu mnVT 
         Caption         =   "T�i kho�n &doanh thu"
         Index           =   11
         Tag             =   "Turnover Account of Finished Contructions"
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnVT 
         Caption         =   "��t/B� TK theo d�i chi ti�t"
         Index           =   13
         Tag             =   "Set Account"
      End
      Begin VB.Menu mnVT 
         Caption         =   "-"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu mnVT 
         Caption         =   "Danh �i�m v�t t�, h�ng ho�"
         Index           =   15
      End
      Begin VB.Menu mnVT 
         Caption         =   "Gi� v�n h�ng nh�p kh�u"
         Index           =   16
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnCongno 
      Caption         =   "C�n&g n�"
      Tag             =   "&Liability"
      Begin VB.Menu mnCN 
         Caption         =   "&Ph�n lo�i"
         Index           =   0
         Tag             =   "Classification..."
      End
      Begin VB.Menu mnCN 
         Caption         =   "&Danh s�ch"
         Index           =   1
         Tag             =   "Items"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnCN 
         Caption         =   "&S� d� ��u k�"
         Index           =   3
         Tag             =   "Opening Balance"
      End
      Begin VB.Menu mnCN 
         Caption         =   "Danh s�ch &H�p ��ng"
         Index           =   4
         Tag             =   "Contract List"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnCN 
         Caption         =   "Ph�n lo�i &nh�n vi�n b�n h�ng"
         Index           =   6
         Tag             =   "Salesman Classification"
      End
      Begin VB.Menu mnCN 
         Caption         =   "Danh s�ch nh�n &vi�n b�n h�ng"
         Index           =   7
         Tag             =   "Salesman List"
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnCN 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnCN 
         Caption         =   "��t/B� TK theo d�i chi ti�t"
         Index           =   11
         Tag             =   "Set Account"
      End
   End
   Begin VB.Menu mnTSCD 
      Caption         =   "T�i &s�n c� ��nh"
      Tag             =   "Fixed &Assets"
      Begin VB.Menu mnTS 
         Caption         =   "Ph�n lo�i &t�i s�n..."
         Index           =   0
         Tag             =   "Classification of Assets..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "Ph�n lo�i &ch�ng t�..."
         Index           =   1
         Tag             =   "Classification of Voucher..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "Danh s�ch TSCD..."
         Index           =   11
         Tag             =   "Classification of Voucher..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnTS 
         Caption         =   "&N��c s�n xu�t..."
         Index           =   3
         Tag             =   "Country List..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "T�nh tr�ng &s� d�ng..."
         Index           =   4
         Tag             =   "Conjuncture List..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "��i t��ng &qu�n l�..."
         Index           =   5
         Tag             =   "Administrative Object..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnTS 
         Caption         =   "T�i s�n ��u &k�..."
         Index           =   7
         Tag             =   "Opening Balance..."
      End
      Begin VB.Menu mnTS 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnTS 
         Caption         =   "��t/B� TK chi ph� kh�u hao"
         Index           =   9
         Tag             =   "Set Account"
      End
      Begin VB.Menu mnTS 
         Caption         =   "Danh s�ch t�i s�n c� ��nh"
         Index           =   10
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Tr� &gi�p"
      Tag             =   "&Help"
      Begin VB.Menu mnuHLP 
         Caption         =   "&N�i dung..."
         Index           =   0
         Tag             =   "&Content..."
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "&Tra c�u..."
         Index           =   1
         Tag             =   "&Search..."
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuHLP 
         Caption         =   "&B�n quy�n"
         Index           =   3
         Tag             =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ret
Dim m_nonClientMetrics As NONCLIENTMETRICS
Dim m_logFont As LOGFONT

Dim m_fontCaption As String * 32

Dim m_fontSmCaption As String * 32
Dim m_fontMenu As String * 32
Dim m_fontMessage As String * 32
Dim m_fontStatus As String * 32
Dim m_fontIcon As String * 32
Dim pProcessEnable As Boolean

Private Const MaxNamTC = 9
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function LDBUser_GetUsers Lib "MSLDBUSR.DLL" (lpszUserBuffer() As String, ByVal lpszFilename As String, ByVal nOptions As Long) As Integer

Private Const OptLDBLoggedUsers = &H2
'Sub chutieng_viet()
'Dim MyUnicodeText
' Set MyUnicodeText = New Class1
'        ' Read Unicode Text from file txtFileName and display in TextBox1(0)
'       'TextBox1(0).Text = MyUnicodeText.ReadUnicode(txtFileName)
'         UVowels = mDOMVowels.ReadUnicode(GetLocalDirectory & "UnicodeVowels.xml")
'       LbCongty.Caption = MyUnicodeText.ReadUnicode("D:\soft\sv\Accounting\config.xml")
'End Sub

Private Sub Command_Click(Index As Integer)
    Select Case Index
        Case 0:
            If User_Right = 2 Then
                NoRight 0
                Exit Sub
            End If
            If pSTOP = 1 Then
                MsgBox VString(pTenCty), vbCritical, App.ProductName
                Exit Sub
            End If
            pPhieu = 0
            FrmChungtu.Show 1
            Set FrmChungtu = Nothing
          Case 1:
            If User_Right = 1 Then
                NoRight 0
                Exit Sub
            End If
            FBcKt.Show 1
         Case 2:
            If User_Right = 1 Then
                NoRight 0
                Exit Sub
            End If
            FBcTC.Show 1
        Case 3:
            RunCT
        Case 4:
            If User_Right = 2 Then
                NoRight 0
                Exit Sub
            End If
            pPhieu = 1
            FrmChungtu.Show 1
            Set FrmChungtu = Nothing
        Case 5:
            If User_Right = 1 Then
                NoRight 0
                Exit Sub
            End If
            FrmBCQT.Show 1
        Case 6:
            FrmCDT.Show 1
    End Select
    HienThongBao "", 1
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CTTimer_Timer()
    If pProcessEnable Then
        pProcessEnable = False
        XuLyChungtu
        pProcessEnable = True
    End If
End Sub

Private Sub email_Click(Index As Integer)
    Select Case Index
        Case 0:
            ShellExecute hwnd, "open", "mailto:" + email(Index).Caption, vbNullString, vbNullString, 0
        Case 1:
            ShellExecute ByVal 0&, "open", email(Index).Caption, vbNullString, vbNullString, 3
        Case 2:
            ShellExecute hwnd, "open", "ypager.exe", vbNullString, Left(pWinDir, 2) + "\Program Files\Yahoo!\Messenger", 1
    End Select
End Sub

Private Sub File1_Click()
'hung
End Sub

Private Sub Form_Activate() ' viet menu

mnDuLieu.Caption = "D� Li�u"

    StationList
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyN:
                RFocus Command(0)
                Command_Click 0
            Case vbKeyC:
                RFocus Command(1)
                Command_Click 1
            Case vbKeyT:
                RFocus Command(2)
                Command_Click 2
        End Select
    End If
    
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyQ Then XoaQuery
        
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyD Then
        ChDir pCurDir + "DATA"
        Recycle "K*" + "_" + CStr(lbCty(0).tag) + ".SAS"
    End If
    
    If (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyF Then
        FontSetUp
        pKhongDau = 1 - pKhongDau
        SetFont Me
        If pKhongDau = 1 Then
             Label(14).Caption = ABCtoKDau(Label(14).Caption)
             Label(26).Caption = ABCtoKDau(Label(26).Caption)
        End If
    End If
    
    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyR Then
        If MsgBox("Xo� t�t c� Relations?", vbYesNo + vbCritical, App.ProductName) = vbYes Then DeleteRel
    End If
    
    If (Shift And vbCtrlMask) > 0 And img.Picture <> 0 And pVersion = 1 Then
        Select Case KeyCode
            Case vbKeyLeft: img.Left = img.Left - 10
            Case vbKeyRight: img.Left = img.Left + 10
            Case vbKeyUp: img.Top = img.Top - 10
            Case vbKeyDown: img.Top = img.Top + 10
        End Select
    End If
    
    If (Shift And vbShiftMask) > 0 And img.Picture <> 0 And pVersion = 1 Then
        Select Case KeyCode
            Case vbKeyLeft: img.Width = img.Width - 10
            Case vbKeyRight: img.Width = img.Width + 10
            Case vbKeyUp: img.Height = img.Height - 10
            Case vbKeyDown: img.Height = img.Height + 10
        End Select
    End If
    
    If (Shift And vbAltMask) > 0 And (Shift And vbCtrlMask) > 0 And KeyCode = vbKeyO Then
            SetPsw pDataPath, pPSW, "unlock"
            WSpace.Close
            End
    End If
End Sub

Private Sub Form_Load()
    Dim X1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer
                
    If findwindowpartial("Microsoft Word") = 0 And findwindowpartial("Microsoft Excel") = 0 Then
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        DoEvents
 
        m_nonClientMetrics.cbSize = Len(m_nonClientMetrics)
        ret = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)
        ret = SystemParametersInfo(SPI_GETICONTITLELOGFONT, Len(m_logFont), m_logFont, 0)
    
        m_fontCaption = m_nonClientMetrics.lfCaptionFont.lfFaceName
        m_fontSmCaption = m_nonClientMetrics.lfSmCaptionFont.lfFaceName
        m_fontMenu = m_nonClientMetrics.lfMenuFont.lfFaceName
        m_fontMessage = m_nonClientMetrics.lfMessageFont.lfFaceName
        m_fontStatus = m_nonClientMetrics.lfStatusFont.lfFaceName
        m_fontIcon = m_logFont.lfFaceName
    
        m_nonClientMetrics.lfCaptionFont.lfFaceName = sFONTNAME & vbNullChar
        m_nonClientMetrics.lfSmCaptionFont.lfFaceName = sFONTNAME & vbNullChar
        m_nonClientMetrics.lfMenuFont.lfFaceName = sFONTNAME & vbNullChar
        m_nonClientMetrics.lfStatusFont.lfFaceName = sFONTNAME & vbNullChar
        m_nonClientMetrics.lfMessageFont.lfFaceName = sFONTNAME & vbNullChar
    
        ret = SystemParametersInfo(SPI_SETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)
        m_logFont.lfFaceName = sFONTNAME & vbNullChar
        ret = SystemParametersInfo(SPI_SETICONTITLELOGFONT, Len(m_logFont), m_logFont, 0)
    End If
   
    If pVersion = 2 Then
        Label(19).Visible = False
      '  img.Top = 3120
      '  img.Left = 360
    End If
    
    If pVersion > 0 Then
        If Len(Dir(pCurDir + "LOGO.JPG")) > 0 Then
            On Error Resume Next
            Set img.Picture = LoadPicture(pCurDir + "LOGO.JPG")
            On Error GoTo 0
            If img.Picture <> 0 Then
                If pVersion <> 2 Then
                    X1 = GetSetting(IniPath, "Logo", "X1", 0)
                    y1 = GetSetting(IniPath, "Logo", "Y1", 0)
                    If X1 <> 0 Then
                        img.Left = X1
                        img.Top = y1
                    End If
                    x2 = GetSetting(IniPath, "Logo", "X2", 0)
                    y2 = GetSetting(IniPath, "Logo", "Y2", 0)
                    If x2 <> 0 Then
                        img.Width = x2
                        img.Height = y2
                    End If
                End If
                img.Visible = True
            End If
        End If
    End If
            
    HienThongBao pDataPath, 2
  '  dlgCommonDialog.InitDir = pCurDir + "DATA"
    
    
    GetLicense
    
    LietKeTep
    
    On Error Resume Next
   ' Rpt.WindowShowPrintSetupBtn = True
 '   Rpt.WindowShowGroupTree = True
  '  Rpt.WindowShowSearchBtn = True
   ' Rpt.WindowShowZoomCtl = True
    On Error GoTo 0
    
    setMDSettings
    
    Select Case pProcessMode
        Case 2: pProcessEnable = True
                        Me.Caption = Me.Caption + " - SERVER Application"
                        CTTimer.Enabled = True
        Case 1: pProcessEnable = False
                        Me.Caption = Me.Caption + " - CLIENT Application"
    End Select
    
    Mask_D = GetShortDateFormat
    
    'chutieng_viet
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
    'kiem tra he thong tai khoan truoc
    KiemTraTaiKhoan
    Screen.MousePointer = 11
    HienThongBao "K�t th�c ch��ng tr�nh k� to�n!", 1
            
    CloseUp
    Recycle pCurDir + "*.BMP"
    
    If pVersion > 0 And img.Visible Then
        SaveSetting IniPath, "Logo", "X1", img.Left
        SaveSetting IniPath, "Logo", "Y1", img.Top
        SaveSetting IniPath, "Logo", "X2", img.Width
        SaveSetting IniPath, "Logo", "Y2", img.Height
    End If
     
    
    m_nonClientMetrics.lfCaptionFont.lfFaceName = m_fontCaption
    m_nonClientMetrics.lfCaptionFont.lfFaceName = m_fontSmCaption
    m_nonClientMetrics.lfMenuFont.lfFaceName = m_fontMenu
    m_nonClientMetrics.lfMessageFont.lfFaceName = m_fontMessage
    m_nonClientMetrics.lfStatusFont.lfFaceName = m_fontStatus

    ret = SystemParametersInfo(SPI_SETNONCLIENTMETRICS, Len(m_nonClientMetrics), m_nonClientMetrics, 0)
    m_logFont.lfFaceName = m_fontIcon
    ret = SystemParametersInfo(SPI_SETICONTITLELOGFONT, Len(m_logFont), m_logFont, 0)
    
    Recycle pCurDir + "DATA\ZIP.MDB"
    
    SetPsw pDataPath, pPSW, ""
    On Error Resume Next
    DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\ZIP.MDB"
   
    On Error GoTo 0
    If Len(Dir(pCurDir + "DATA\ZIP.MDB")) > 0 Then
        Recycle pDataPath
        FileCopy pCurDir + "DATA\ZIP.MDB", pDataPath
        SetPsw pCurDir + "DATA\ZIP.MDB", "", pPSW
    End If
    SetPsw pDataPath, "", pPSW
    '========================
     Recycle pCurDir + "DATA\AJZIP.MDB"
    
    SetPsw pDataPath, pPSW, ""
    On Error Resume Next
    DBEngine.CompactDatabase pDataPath, pCurDir + "DATA\AJZIP.MDB"
   
    On Error GoTo 0
    If Len(Dir(pCurDir + "DATA\AJZIP.MDB")) > 0 Then
        Recycle pDataPath
        FileCopy pCurDir + "DATA\AJZIP.MDB", pDataPath
        SetPsw pCurDir + "DATA\AJZIP.MDB", "", pPSW
    End If
    SetPsw pDataPath, "", pPSW
    '========================
    restoreSettings
   
    Screen.MousePointer = 0
        
    End
    Set App = Nothing
End Sub


Private Sub Image1_Click()

End Sub

Public Sub mnCn_Click(Index As Integer)
    If Index = 3 Or Index = 9 Then
        If Not KtraMKAdmin Then Exit Sub
    End If
    Select Case Index
        Case 0:
            frmPhanLoaiVT.tag = 2
            frmPhanLoaiVT.Show 1
        Case 1:
            FrmKhachHang.Show vbModal
        Case 3:
            If ChoDieuChinhDauKy Then
                If pCongNoHD = 0 Then
                    FKHDauKy.Show vbModal
                Else
                    FKHDauKy2.Show vbModal
                End If
            End If
        Case 4:
            FrmHD.Show vbModal
        Case 6:
            frmPhanLoaiVT.tag = 4
            frmPhanLoaiVT.Show 1
        Case 7:
            FrmNhanVien.Show 1
        Case 9:
            If KtraMKAdmin Then FrmLS.Show 1
        Case 11:
            If KtraMKAdmin Then DatTKCN
    End Select
    HienThongBao "", 1
End Sub

Private Sub mnDL_Click(Index As Integer)
    Dim sql As String
    
    If User_Right <> 0 Or (Me.tag Mod 10 = 0) Or (User_Right = 2) Then
        NoRight 0
        Exit Sub
    End If
    Me.MousePointer = 11
    
     
    Select Case Index
        Case 0:
            If Not STDetail Then
                NoRight 1
                GoTo KT
            End If
            KiemTraVatTu
         Case 20:
           
               FrmNguyente.Show 1
          
        Case 1:
            KiemTraTaiKhoan
        Case 3:
            If FPsw.GetPswX() = "UCDIT" Then
                sql = FrmGetStr.GetString("L�nh x� l�:", App.ProductName)
                If Len(sql) > 0 Then ExecuteSQL5 sql
            End If
        Case 6:
            If KtraMKAdmin Then
                If MsgBox("B�n ch�c ch�n k�t th�c n�m " + CStr(pNamTC) + " v� chuy�n sang n�m m�i ?" _
                    , vbYesNo + vbExclamation, App.ProductName) <> vbYes Then GoTo KT
                
                HienThongBao "Chuy�n s� d� cu�i k� ...  Xin vui l�ng ch� !", 1
                ChuyenNamMoi
                lbCty(7).Caption = CStr(pNamTC)
                LietKeNam
            End If
        Case 9: FrmKC.Show vbModal
        Case 10: FrmPBCP.Show vbModal
        Case 11: FrmThKC.Show vbModal
        Case 14: FrmCTGS.Show vbModal
        Case 16:
            sql = GetSetting(IniPath, "LastYear", "IncTax" + CStr(pNamTC), 0)
            sql = InputBox("S� �i�u ch�nh", "Thu� thu nh�p doanh nghi�p " + CStr(pNamTC - 1), sql)
            If IsNumeric(sql) Then SaveSetting IniPath, "LastYear", "IncTax" + CStr(pNamTC), sql
        Case 17:
            sql = ChonTenTep("Ch�n t�p d� li�u c�a n�m TC tr��c (L�u � c�n ch�y ki�m tra s� li�u c�a n�m c�)", &H4&, "*.MDB", 1)
            If Len(sql) = 0 Then GoTo KT
            LaySoDauNam sql
        Case 19: If KtraMKAdmin Then FrmE.Show 1
        Case 21:
            If KtraMKAdmin Then
                sql = FrmDB.ChonTepLuu(frmMain.lbCty(8).Caption, pNamTC)
                If Len(sql) > 0 Then
                    CloseUp 1
                    OpenDB sql
                End If
            End If
    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Public Sub mnHT_Click(Index As Integer)
    Dim psw As String, st As Integer, fn As String
    
    If Index = 5 Or Index = 6 Or Index = 10 Then
        If Not KtraMKAdmin Then Exit Sub
    End If
    Me.MousePointer = 11
    Select Case Index
        Case 0:                                             ' Mo tep
a:
            psw = ChonTenTep("Ch�n t�p d� li�u", &H4&, "*.MDB", 1)
MoTep:
            If Len(psw) = 0 Then GoTo KT
            HienThongBao "M� t�p d� li�u...", 1
            If st = 0 Then CloseUp 1
            If OpenDB(psw, 1) = 0 Then
                GetLicense
                
                If pDataPath <> GetSetting(IniPath, "Environment", "Path") Then
                    pProcessMode = 0
                Else
                    Select Case UCase(App.EXEName)
                        Case "SERVER":  pProcessMode = 2
                        Case "CLIENT":  pProcessMode = 1
                        Case Else: pProcessMode = 0
                    End Select
                End If
                
                FrmMatkhau.Show 1
                Set FrmMatkhau = Nothing
                SetUserRight
                
                LietKeTep
            Else
                st = 1
                GoTo a
            End If
        Case 1:                                             ' Sao chep
            DelTemp
            psw = ChonTenTep("Sao ch�p t�p d� li�u", &H4&, "*.MDB", 2)
            If Len(psw) = 0 Then GoTo KT
            Me.MousePointer = 11
            HienThongBao "Sao ch�p t�p d� li�u ...", 1
            CloseUp 1
            On Error Resume Next
            DBEngine.CompactDatabase pDataPath, psw, , , ";pwd=" + pPSW
            On Error GoTo 0
            OpenDB pDataPath
        Case 2:                                             ' Tep mac dinh
            mnHT_Click 0
            SaveSetting IniPath, "Environment", "Path", pDataPath
        Case 3:                                             ' Nen tep du lieu
            DelTemp
            psw = ChonTenTep("N�n t�p d� li�u", &H4&, "*.SAS", 2)
            If Len(psw) = 0 Then GoTo KT
            Me.MousePointer = 11
            HienThongBao "N�n t�p d� li�u ...", 1
            CloseUp 1
            Recycle pCurDir + "TEMP.MDB"
            On Error Resume Next
            DBEngine.CompactDatabase pDataPath, pCurDir + "TEMP.MDB", , , ";pwd=" + pPSW
            On Error GoTo 0
            If Len(Dir(pCurDir + "TEMP.MDB")) > 0 Then
                NenTep pCurDir + "TEMP.MDB", psw
                Recycle pCurDir + "TEMP.MDB"
            Else
                NenTep pDataPath, psw
            End If
X1:
            OpenDB pDataPath
        Case 4:
            psw = ChonTenTep("Ch�n t�p d� li�u n�n", &H4&, "*.SAS", 1)
            If Len(psw) = 0 Then GoTo KT
            fn = ChonTenTep("Ch�n t�n t�p d� li�u", &H4&, "*.MDB", 2)
            If Len(fn) = 0 Then GoTo KT
            GianTepNen psw, fn
            
            CloseUp 1
                       
            OpenDB fn, 1
            GetLicense
            
            FrmMatkhau.Show 1
            Set FrmMatkhau = Nothing
            SetUserRight
        Case 5:
            EMailDB
        Case 6:
            psw = ChonTenTep("T�n t�p d� li�u", &H4&, "*.MDB", 2)
            If Len(psw) = 0 Then GoTo KT
            CloseUp 1
            On Error GoTo KT
            DBEngine.CompactDatabase pDataPath, psw, , , ";pwd=" + pPSW
            On Error GoTo 0
            OpenDB psw
            ExecuteSQL5 "UPDATE License SET LoaiTien=" + IIf(pTien = 0, "1", "0")
            GetLicense
            DoiTyGiaDB
        Case 8:                                             ' Dat may in
            ChonTenTep "", 0, "", 3
        Case 9:                                             ' Dat may in
            ChonTenTep "", cdlCFBoth, "", 4
            If Len(dlgCommonDialog.FontName) > 1 And (LoaiFont(dlgCommonDialog.FontName) = FontFlag Or KiemTraMaSoThue(lbCty(8).Caption, "03")) Then
                pFontName = dlgCommonDialog.FontName
                pFontSize = dlgCommonDialog.FontSize
                ExecuteSQL5 "UPDATE License SET FontName='" + pFontName + "', FontSize=" + CStr(pFontSize)
                lbCty(0).FontName = pFontName
                lbCty(1).FontName = pFontName
                mnHT(10).Caption = IIf(FontFlag <> 2, "Chuy�n ��i CSDL sang font ABC", "Chuy�n ��i CSDL sang font VNI")
                SetFont Me
            End If
        Case 10:
            If MsgBox("B�n ch�c ch�n c�n ��i font ? (Ch� � ch�n font ch� tr��c khi ��i)", vbCritical + vbYesNo, App.ProductName) = vbNo Then GoTo KT
            Me.MousePointer = 11
            ChuyenDoiFont FontFlag = 2
            GetLicense
        Case 11:                                             ' Thong so
            If User_Right = 0 Then
                FrmOptions.Show 1
                GetLicense
            Else
                NoRight 0
            End If
        Case 13:                                             ' Danh sach user
            If User_Right = 0 Then
                FrmUser.Show 1
            Else
                NoRight 0
            End If
        Case 14:                                           ' Dat mat khau
            'Load FrmMatkhau
            FrmMatkhau.tag = 1
            FrmMatkhau.Show 1
        Case 16:
            If (Not IsNumeric(Left(lbCty(8).Caption, 2))) Then GoTo KT
            If CInt(Left(lbCty(8).Caption, 3)) = 0 Then GoTo KT
            If (Len(pMST) > 0 And Left(lbCty(8).Caption, Len(pMST)) = pMST) Then GoTo B
            If FrmGetStr.GetMK(lbCty(8).Caption) Then
B:
                UpDateDB
                GetLicense
            End If
        Case 18 To 22:
            psw = mnHT(Index).Caption
            GoTo MoTep
        Case 24:
            FrmMatkhau.Show 1
            SetUserRight
        Case 25:
            Unload Me
            Exit Sub
    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnK_Click(Index As Integer)
    Dim k As Integer
    
    If User_Right <> 0 Then
        NoRight 0
        Exit Sub
    End If
    
    k = SelectSQL("SELECT Lock" + CStr(mnk(Index).tag) + " Mod 10 AS F1 FROM License")
    If MsgBox("C�n " + IIf(k = 0, "kho�", "cho nh�p") + IIf(mnk(Index).tag > 0, " ph�t sinh th�ng " + CStr(mnk(Index).tag), " s� d� ��u n�m") + " ?", vbYesNo + vbExclamation, App.ProductName) <> vbYes Then Exit Sub
    ExecuteSQL5 "UPDATE License SET Lock" + CStr(mnk(Index).tag) + "=10*(Lock" + CStr(mnk(Index).tag) + " \ 10)+" + CStr(1 - k)
    mnk(Index).Caption = IIf(1 - k > 0, Trim(mnk(Index).Caption) + "          x", Left(mnk(Index).Caption, Len(mnk(Index).Caption) - 1))
End Sub

Private Sub mnkt_Click(Index As Integer)

End Sub

Private Sub mnNam_Click(Index As Integer)
    Dim i As Integer, path As String
    
    Me.MousePointer = 11
    CloseUp 1
    If Index = 4 Then
        path = GetSetting(IniPath, "Environment", "Path", pCurDir + "DATA\KETOAN.MDB")
    Else
        path = GetSetting(IniPath, "LastYear", mnNam(Index).Caption, pCurDir + "DATA\KETOAN.MDB")
    End If
    If OpenDB(path) <> 0 Then mnHT_Click 0
    For i = 0 To 4
        mnNam(i).CHECKED = (i = Index)
    Next
    pNamTC = CInt5(mnNam(Index).Caption)
    
    lbCty(7).Caption = CStr(pNamTC)
    Me.MousePointer = 0
End Sub

Private Sub mnnh_Click(Index As Integer)

End Sub

Public Sub mnTS_Click(Index As Integer)
    If (Not FADetail) Or User_Right = 2 Then
        NoRight 2
        Exit Sub
    End If
    Me.MousePointer = 11
      
    Select Case Index
        Case 0:                         ' Phan loai TS
            'Load frmPhanLoai
            frmPhanLoai.tag = 1
            frmPhanLoai.Show 1
        Case 11:                         ' Phan loai TS
          frmDSTaiSan.Show 1
        
        Case 1:                         ' Phan loai ctu
            'Load frmPhanLoai
            frmPhanLoai.tag = 2
            frmPhanLoai.Show 1
        Case 3:                         ' Nuoc sx
            'Load FrmKho
            FrmKho.tag = 2
            FrmKho.Show 1
        Case 4:                         ' Tinh trang SD
            'Load FrmKho
            FrmKho.tag = 3
            FrmKho.Show 1
        Case 5:                         ' DTQL
            'Load FrmKho
            FrmKho.tag = 4
            FrmKho.Show 1
        Case 7:
            If ChoDieuChinhDauKy Then
                pNghiepVu = NV_TANG
                'Load frmTaiSan
                frmTaiSan.tag = 1
                frmTaiSan.Show 1
            End If
        Case 9:
            If KtraMKAdmin Then DatTKTS
        Case 10:
             frmDSTaiSan.Show 1
    End Select
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnuHLP_Click(Index As Integer)
    Dim nRet As Integer
    
    Select Case Index
        Case 0:                                             ' Noi dung
            On Error Resume Next
            nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
            If Err Then MsgBox Err.Description
            On Error GoTo 0
        Case 1:                                             ' Tra cuu
            On Error Resume Next
            nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
            If Err Then MsgBox Err.Description
            On Error GoTo 0
        Case 3:                                             ' Ban quyen
            frmAbout.Show vbModal, Me
    End Select
End Sub

Private Sub mnunh_Click()

End Sub

Private Sub mnviet_Click()

End Sub

Public Sub mnVT_Click(Index As Integer)
    Dim st As String, i As Integer, TK As String, d1 As Date, d2 As Date, j As Integer, k As Integer, mv As Long

    If User_Right = 2 Then
        NoRight 0
        Exit Sub
    End If
    
    If Not STDetail Then
        NoRight 1
        Exit Sub
    End If
    
    If Index = 3 Or Index = 4 Or Index = 5 Or Index = 10 Or Index = 12 Then
        If Not KtraMKAdmin Then Exit Sub
    End If
    
    Me.MousePointer = 11
    
     
    Select Case Index
        Case 0:
            frmPhanLoaiVT.tag = 1
            frmPhanLoaiVT.Show 1
       Case 20:
            FrmKho.tag = 1
            FrmKho.Show 1
       Case 17:
            FrmVattu.Show 1
                
        Case 18:
          FrmLuuChuyen.Show 1
          Case 19:
             FrmThanhPham.Show 1
        Case 1:
            FrmNguon.Show 1
        Case 3:
            If ChoDieuChinhDauKy Then FVTDauKy.Show 1
        Case 4:
            If OutCost <> 2 Then
                st = FrmGetStr.GetString("Th�ng c�n t�nh l�i:", "T�nh gi� xu�t kho")
                If IsNumeric(st) Then
                    i = CInt5(st)
                    j = i
                Else
                    i = InStr(st, "-")
                    If i > 0 Then
                        j = CInt5(Right(st, Len(st) - i))
                        i = CInt5(Left(st, i - 1))
                    Else
                        i = CInt5(st)
                        j = i
                    End If
                End If
            Else
                i = 1
                j = 12
            End If
            If i > 0 And i < 13 And j > 0 And j < 13 Then
                st = ""
                st = FrmGetStr.GetString("S� hi�u v�t t� c�n t�nh l�i (�� tr�ng n�u t�nh l�i to�n b�):", "T�nh gi� xu�t kho")
                Do While Len(st) > 0
                    mv = SoHieu2MaSo(st, "Vattu")
                    If mv > 0 Then Exit Do
                    st = FrmGetStr.GetString("S� hi�u v�t t� c�n t�nh l�i (�� tr�ng n�u t�nh l�i to�n b�):", "T�nh gi� xu�t kho")
                Loop
                If OutCost <> 2 Then TK = FrmGetStr.GetString("S� hi�u t�i kho�n ghi n� khi xu�t kho c�n t�nh l�i (�� tr�ng n�u t�nh l�i to�n b�):", "T�nh gi� xu�t kho", "") Else TK = ""
                Me.MousePointer = 11
                If OutCost = 0 Then
                    k = CInt5(FrmGetStr.GetString("Nh�p s� 1 �� t�nh b�nh qu�n di ��ng, s� 2 �� t�nh b�nh qu�n cu�i k� ", "T�nh l�i gi� xu�t kho"))
                    If k < 1 And k > 2 Then GoTo KT
                    If k = 1 Then TinhGXK i, j, st, TK
                    If k = 2 Then TinhGXKBQ i, j, st, TK
                End If
                If OutCost = 1 Then TinhGVBH NgayDauThang(pNamTC, pThangDauKy), NgayCuoiNam(), 1, mv
                If OutCost = 2 Then TinhGXKFIFO i, j, st, TK
            End If
        Case 5:
            If OutCost = 2 Then
                d1 = NgayDauThang(pNamTC, pThangDauKy)
                d2 = NgayCuoiNam
            Else
                If Not GetDate2.GetDate("T�nh gi� v�n b�n h�ng", d1, d2) Then Exit Sub
            End If
            
            i = MsgBox("L�p l�i c�c ch�ng t� gi� v�n �� t�nh ? (N�u kh�ng th� ch��ng tr�nh ch� l�p c�c ch�ng t� gi� v�n c�n thi�u)", vbCritical + vbYesNo, App.ProductName)
            st = FrmGetStr.GetString("S� hi�u v�t t� c�n t�nh l�i (�� tr�ng n�u t�nh l�i to�n b�):", "T�nh gi� v�n")
            Do While Len(st) > 0
                mv = SoHieu2MaSo(st, "Vattu")
                If mv > 0 Then Exit Do
                st = FrmGetStr.GetString("S� hi�u v�t t� c�n t�nh l�i (�� tr�ng n�u t�nh l�i to�n b�):", "T�nh gi� v�n")
            Loop
                
            Me.MousePointer = 11
            If OutCost = 0 Then
                k = CInt5(FrmGetStr.GetString("Nh�p s� 1 �� t�nh b�nh qu�n di ��ng, s� 2 �� t�nh b�nh qu�n cu�i k� (t�nh theo th�ng)", "T�nh l�i gi� v�n"))
                If k < 1 And k > 2 Then GoTo KT
                TinhGVBHBQ Month(d1), Month(d2), i, mv, k
            Else
                TinhGVBH d1, d2, i, mv
            End If
        Case 6:
            FVTDauKy.tag = 1
            FVTDauKy.Show 1
        Case 7:
            KiemKeN.Show vbModal
        Case 9:
            'Load frmPhanLoaiVT
            frmPhanLoaiVT.tag = 3
            frmPhanLoaiVT.Show 1
        Case 10:
            FrmTP.Show 1
        Case 11:
            If KtraMKAdmin Then DatTKDTTP
        Case 13:
            If KtraMKAdmin Then DatTKVT
        Case 15 ', 16:
             FrmVattu.Show 1
           ' CPGV.tag = Index - 15
            'CPGV.Show 1
    End Select
KT:
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub mnXoa_Click(Index As Integer)
    
    If User_Right <> 0 Then
        NoRight 0
        Exit Sub
    End If
    
    
    If mnXoa(Index).tag > 0 Then
        If MsgBox("B�n ch�c ch�n c�n x�a ph�t sinh th�ng " + CStr(mnXoa(Index).tag) + " ?", vbYesNo + vbExclamation, App.ProductName) = vbYes Then
            Me.MousePointer = 11
            HienThongBao "X�a ph�t sinh th�ng " + CStr(mnXoa(Index).tag) + " ...  Xin vui l�ng ch� !", 1
            XoaPSThang Index
        End If
    Else
        If MsgBox("B�n ch�c ch�n c�n x�a s� d� ��u n�m?", vbYesNo + vbExclamation, App.ProductName) = vbYes Then
            Me.MousePointer = 11
            HienThongBao "X�a s� d� ��u n�m, xin vui l�ng ch� !", 1
            XoaDK
        End If
    End If
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

Private Sub OptNN_Click(Index As Integer)
    CloseItemList
    pNN = Index
'    Img.Visible = (pNN = 0)
    SetFont Me, 1
End Sub

Private Sub sbStatusBar_PanelClick(ByVal Panel As ComctlLib.Panel)
    Dim path As String
    
    Select Case Panel.Index
        Case 2:             Panel.Text = IIf(Panel.tag = 0, "Data File Size: " + Format(FileLen(pDataPath) / 1048576, Mask_2) + " MB, Version: " + IIf(DBKetoan.Version < 4, "97", "2000"), pDataPath)
                                    Panel.tag = 1 - Panel.tag
                                    StationList
        Case 3:             Panel.Text = IIf(Panel.tag = 0, Panel.ToolTipText, UserName)
                                    Panel.tag = 1 - Panel.tag
        Case 4:             path = GetSetting(IniPath, "Environment", "BackUpPath")
                                    path = FrmGetStr.GetString("Th� m�c l�u d� li�u", App.ProductName, path)
                                    SaveSetting IniPath, "Environment", "BackUpPath", path
    End Select
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Me.MousePointer = 11
    
    Select Case Button.key
        Case "TaiKhoan"
            FrmTaikhoan.tag = 1
            FrmTaikhoan.Show 1
        Case "NgoaiTe"
            FrmNguyente.Show 1
        Case "Kho"
            If STDetail Then
                'Load FrmKho
                FrmKho.tag = 1
                FrmKho.Show 1
            Else
                NoRight 1
            End If
        Case "VatTu"
            If STDetail Then
                FrmVattu.Show 1
            Else
                NoRight 1
            End If
        Case "LuuChuyen"
            If STDetail Then
                FrmLuuChuyen.Show 1
            Else
                NoRight 1
            End If
        Case "DuPhong"
            If STDetail Then
                FrmDuphong.Show 1
            Else
                NoRight 1
            End If
        Case "TaiSan"
            If FADetail Then
                pNghiepVu = NV_KHONG
                frmDSTaiSan.Show 1
            Else
                NoRight 2
            End If
        Case "CN"
            If KHDetail Then FrmKhachHang.Show vbModal
        Case "TongHop"
            FrmTongHop.Show 1
        Case "Help"
            mnuHLP_Click 0
        Case "KetThuc"
         'FVAT.tag = 1
         ' FVAT.Show 1
            Unload Me
            Exit Sub
        Case "ThanhPham"
            If STDetail Then
                FrmThanhPham.Show 1
            Else
                NoRight 1
            End If
    End Select
    HienThongBao "", 1
    Me.MousePointer = 0
End Sub

'======================================================================================
' Function GetLicense : Th� t�c l�y t�n c�ng ty v� chi nh�nh
'======================================================================================
Private Sub GetLicense()
    Dim rs_license As Recordset, i As Integer, k As Integer, sh As String
    
    CloseItemList
    DEMO = 1
    i = er_DBFile
    On Error Resume Next
    Set rs_license = DBKetoan.OpenRecordset("License", dbOpenSnapshot)
    i = 0
    i = rs_license!Flag1 Mod 100
    If i > 0 Then
        If StationList() > i Then
            i = er_Connection
            Unload Me
            Exit Sub
        End If
    End If
    
    pTenCty = rs_license!TenCty
    pTenCn = rs_license!tencn
    
    lbCty(2).Caption = rs_license!DiaChi
    lbCty(3).Caption = rs_license!Tel
    lbCty(4).Caption = rs_license!Fax
    lbCty(5).Caption = rs_license!TaiKhoanVN
    lbCty(6).Caption = rs_license!TaiKhoanNT
    pNamTC = rs_license!NamTC
    pThangDauKy = rs_license!thang
    lbCty(7).Caption = CStr(pNamTC)
    lbCty(8).Caption = rs_license!MaSoThue
    
    pBaoGia = (rs_license!Flag1 Mod 1000) \ 100
    pNVBH = (rs_license!Flag1 Mod 10000) \ 1000
    
    For i = 5 To 7
        mnCN(i).Visible = (pNVBH > 0)
    Next
    lb(0).tag = "Model"
    SetFont Me
    i = (rs_license!Flag1 Mod 1000000000) \ 100000000
    lb(0).tag = i
    If (i < 3 Or i = 5) And pVersion = 0 Then ExecuteSQL5 "UPDATE License SET Flag1=400000000+Flag1 Mod 100000000", False
    Select Case i
        Case 1: lb(1).Caption = "Doanh nghi�p Nh� n��c"
                        lb(0).Caption = "8.0.NN."
        Case 2: lb(1).Caption = "C� ph�n - Li�n doanh"
                        lb(0).Caption = "8.0.CP."
        Case 3: lb(1).Caption = "C�ng ty TNHH"
                        lb(0).Caption = "8.0.HH."
        Case 4: lb(1).Caption = "Doanh nghi�p t� nh�n"
                        lb(0).Caption = "8.0.TN."
        Case 5: lb(1).Caption = "C� s� ��o t�o"
                        lb(0).Caption = "8.0.DT."
        Case 6:
                        lb(1).Caption = "H�nh ch�nh s� nghi�p"
                        lb(0).Caption = "8.0.HCSN."
                        Label(24).Visible = False
                        Label(25).Visible = False
                        Frame(1).Visible = False
        Case Else
                        lb(0).Caption = "8.00"
    End Select
    If pVersion <> 3 Then lb(0).Caption = lb(0).Caption + IIf((rs_license!Flag1 Mod 100000000) \ 10000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 10000000) \ 1000000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 1000000) \ 100000 > 0, "1", "0") + IIf((rs_license!Flag1 Mod 100000) \ 10000 > 0, "1", "0")
    Chk(0).Value = (rs_license!Flag1 Mod 100000000) \ 10000000
    Chk(1).Value = (rs_license!Flag1 Mod 10000000) \ 1000000
    Chk(2).Value = (rs_license!Flag1 Mod 1000000) \ 100000
    Chk(3).Value = (rs_license!Flag1 Mod 100000) \ 10000
        
    Command(6).Visible = ((rs_license!Flag1 Mod 1000000) \ 100000 > 0)
    
    Command(4).Visible = (rs_license!Flag1 \ 1000000000 > 0)
            
    pTygia = IIf(rs_license!tygia > 0, 1, 0)
    pHachToan = 1 - (rs_license!RptOrder Mod 10)
    pRpt = (rs_license!RptOrder Mod 100) \ 10
    OutCost = rs_license!OutCost
    FCost = rs_license!FixedOutCost
    STDetail = rs_license!STDetail
    FADetail = rs_license!FADetail
    KHDetail = rs_license!HDV
    pGiaHT = rs_license!GiaHT
    pGiaVon = (rs_license!STDetail Mod 100) \ 10
    pDTTP = (rs_license!STDetail Mod 1000) \ 100
    pDinhmuc = (rs_license!STDetail Mod 10000) \ 1000
        
    Command(5).Visible = ((rs_license!Lock0 Mod 100) \ 10 > 0)
    pCongNoHD = (rs_license!Lock0 Mod 1000) \ 100
    pPQTK = (rs_license!Lock0 Mod 10000) \ 1000
    pGiaUSD = (rs_license!Lock0 Mod 100000) \ 10000
    pChietKhau = (rs_license!Lock1 Mod 100) \ 10
    pKiemKeNgay = (rs_license!Lock1 Mod 1000) \ 100
    pNoiBo = (rs_license!Lock1 Mod 10000) \ 1000
    pSoVV = (rs_license!Lock1 Mod 100000) \ 10000
    pNhapKhau = (rs_license!Lock2 Mod 100) \ 10
    pBarCode = (rs_license!Lock2 Mod 1000) \ 100
    pNhapDoiTuong = (rs_license!Lock2 Mod 10000) \ 1000
    pTrungSoHieuKhacThang = (rs_license!Lock2 Mod 100000) \ 10000
    
    mnVT(14).Visible = (pNhapKhau > 0)
    mnVT(15).Visible = (pNhapKhau > 0)
    
    pTien = 0
    pTien = rs_license!loaitien
    If pTien > 0 Then
        Mask_0 = Mask_2
        pTienStr = TenNT(pTien)
    Else
        Mask_0 = GetSetting(IniPath, "Environment", "IntMask", "###,###,###,###")
        pTienStr = "VND"
    End If
    CTGS_GV = rs_license!CTGS_GV
    pFontName = rs_license!FontName
    pFontSize = rs_license!FontSize
    lbCty(0).FontName = pFontName
    lbCty(1).FontName = pFontName
    lbCty(10).Caption = rs_license!Quan
    lbCty(11).Caption = rs_license!ThanhPho
    frmMain.lbCty(9).Caption = rs_license!email
    pSoKT = rs_license!SoKT
    mnDL(13).Visible = (pSoKT Mod 100 >= 10)
    mnDL(14).Visible = (pSoKT Mod 100 >= 10)
    tbToolBar.Buttons("ThanhPham").Visible = (rs_license!RptOrder Mod 10000 >= 1000)
    tbToolBar.Buttons("ThanhPham2").Visible = (rs_license!RptOrder Mod 10000 >= 1000)
    pSongNgu = False
    pSongNgu = (pSoKT Mod 100000 >= 10000)
    pMaVach = 0
    pMaVach = rs_license!mv Mod 10
    pTyGiaBQ = 0
    pTyGiaBQ = IIf(rs_license!mv Mod 10000 >= 1000, 1, 0)
    tbToolBar.Buttons("TongHop").Visible = False
    tbToolBar.Buttons("TongHop").Visible = (rs_license!mv Mod 1000 >= 100)
    DEMO = IIf((rs_license!mv Mod 100000 >= 10000) And (rs_license!MKUP = pRev), 0, 1)
    NgayDauThangMoi = rs_license!NgayDauThang
    FontFlag = LoaiFont(pFontName)

    If (Not pSongNgu) And OptNN(1).Value Then OptNN(0).Value = True
    i = pNN
    pNN = 0
    
    pNN = i
    mnVT(4).Visible = (OutCost = 0 Or OutCost = 1 Or OutCost = 2)
    mnVT(7).Visible = (pKiemKeNgay > 0)
    For i = 8 To 11
        mnVT(i).Visible = (pDTTP <> 0)
    Next
    mnDL(19).Visible = pSongNgu
    
    sh = SelectSQL("SELECT App1Path AS F1 FROM License")
    Command(3).Visible = Len(Dir(sh)) > 0
    
    mnHT(6).Visible = (pTygia > 0)
    mnHT(10).Caption = IIf(FontFlag <> 2, "Chuy�n ��i CSDL sang font ABC", "Chuy�n ��i CSDL sang font VNI")
    mnHT(10).Visible = (rs_license!RptOrder Mod 1000 >= 100)
    
    mnCongno.Visible = KHDetail
  
    Me.Caption = "Sao Viet Accounting Software - "
    sh = LaySH(rs_license!TKVattu, 1, "-")
    If DEMO = 0 And pVersion <> 2 Then
        Me.Caption = Me.Caption + "8.0" '+ sh
                        
        If ((Int_StrToCode(rs_license!MaSoThue) <> rs_license!MST_ID) Or (Int_StrToCode(pTenCty) <> rs_license!TenCty_ID) Or (Int_StrToCode(pTenCn) <> rs_license!tencn_id)) Then
            pTenCty = ABCtoVNI("Sao ch�p kh�ng b�n quy�n")
            pTenCn = ABCtoVNI("Sao ch�p kh�ng b�n quy�n")
            ExecuteSQL5 "UPDATE License SET MST_ID=-1"
            pSTOP = 1
        End If
    Else
        Me.Caption = Me.Caption + sh + IIf(pVersion < 2, " - Training Version", " - Ch��ng tr�nh ph�i h�p ��o t�o")
    End If
    
    If pVersion = 3 Then
        Me.Caption = Me.Caption + " - HCSN"
        pVATV = "3113"
        pSHPT = "3111"
    Else
        pVATV = "133"
        pSHPT = "131"
    End If
    
    lbCty(0).tag = rs_license!TenCty_ID
    lbCty(0).Caption = pTenCty
    lbCty(1).Caption = pTenCn
    Frame(0).Visible = pSongNgu
    
    mnXoa(0).tag = 0
    mnk(0).tag = 0
    mnk(0).Caption = mnk(0).Caption + IIf(rs_license.Fields("Lock0") Mod 10 > 0, "          x", "")
    For i = 1 To 12
        k = CThangFR(i)
        sh = IIf(rs_license.Fields("Lock" + CStr(i)) Mod 10 > 0, "          x", "")
        mnXoa(i).Caption = CStr(k) + "/" + CStr(pNamTC)
        mnk(i).Caption = CStr(k) + "/" + CStr(pNamTC) + sh
        mnXoa(i).tag = k
        mnk(i).tag = k
    Next
            
    rs_license.Close
    Set rs_license = Nothing
    
    LietKeNam
     mnVT(15).Visible = True
     
    On Error GoTo 0
End Sub

Private Sub NoRight(id As Integer)
    Select Case id
        Case 0: HienThongBao "Kh�ng c� quy�n truy c�p!", 1
        Case 1: HienThongBao "Kh�ng ��ng k� theo d�i chi ti�t v�t t�!", 1
        Case 2: HienThongBao "Kh�ng ��ng k� theo d�i chi ti�t TSC�!", 1
    End Select
    Beep
End Sub

Private Sub LietKeNam()
    Dim rs As Recordset, i As Integer
        
    mnNam(MaxNamTC).Caption = CStr(pNamTC)
    mnNam(MaxNamTC).CHECKED = True
    If Not BangDaCo("NamTC") Then Exit Sub
    Set rs = DBKetoan.OpenRecordset("SELECT * FROM NamTC WHERE Nam<" + CStr(pNamTC) + " ORDER BY Nam DESC")
    i = MaxNamTC
    Do While (i > 0) And (Not rs.EOF)
        i = i - 1
        mnNam(i).Caption = CStr(rs!nam)
        mnNam(i).Visible = True
        mnNam(i).tag = rs!path
        rs.MoveNext
    Loop
    Do While (i > 0)
        i = i - 1
        mnNam(i).Visible = False
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Public Sub SetUserRight()
    Dim i As Integer
    
    Command(0).Enabled = (User_Right <> 2)
    Command(4).Enabled = (User_Right <> 2)
    
    For i = 1 To 11
        tbToolBar.Buttons(i).Enabled = (User_Right <> 2)
    Next
    
    For i = 1 To 4
        mnHT(i).Enabled = (User_Right = 0)
    Next
    
    For i = 10 To 11
        mnHT(i).Enabled = (User_Right = 0)
    Next
    
    mnHT(13).Enabled = (User_Right = 0)
    mnDL(0).Enabled = (User_Right = 0)
    mnDL(1).Enabled = (User_Right = 0)
    mnDL(3).Enabled = (User_Right = 0)
    mnDL(4).Enabled = (User_Right = 0)
    mnDL(7).Enabled = (User_Right = 0)
    
    For i = 9 To 12
        mnDL(i).Enabled = (Me.tag Mod 10 >= 1) Or (User_Right = 0)
    Next
    'mnKC(4).Enabled = (User_Right = 0)
    
    mnVatTu.Enabled = (Me.tag Mod 100 >= 10) Or (Me.tag Mod 1000 >= 100)
    mnTSCD.Enabled = (Me.tag Mod 10000 >= 1000)
    mnCongno.Enabled = (Me.tag Mod 100000 >= 10000)
    Command(2).Enabled = (User_Right <> 1) And (Me.tag Mod 10 >= 1)
    Command(6).Enabled = (User_Right <> 1) And (Me.tag Mod 10 >= 1)
End Sub

Private Sub DatTKCN()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 2
    shtk = FrmGetStr.GetString("S� hi�u TK", "��t/B� TK theo d�i chi ti�t")
    If Len(shtk) = 0 Then GoTo KT
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    If TK.tk_id = TKVT_ID Or TK.tk_id = TSCD_ID Or TK.tk_id = KHTSCD_ID Or TK.tk_id = TKThue_ID Or TK.tk_id = TKDT_ID Then Exit Sub
    If TK.TkCoPS(0, 0) Or TK.NoDauKy <> 0 Or TK.CoDauKy <> 0 Then
        Me.MousePointer = 11
        If TK.ChuyenChiTietSangDoiTuong Then
            MsgBox "C�c chi ti�t t�i kho�n �� ���c m� ho� th�nh ��i t��ng c�ng n�!", vbCritical, App.ProductName
        Else
            MsgBox "T�i kho�n kh�ng chuy�n ��i ���c!", vbCritical, App.ProductName
        End If
        Me.MousePointer = 0
        GoTo KT
    End If
    If TK.tk_id = TKCNKH_ID Or TK.tk_id = TKCNPT_ID Then ExecuteSQL5 "DELETE SoDuKhachHang.* FROM SoDuKhachHang INNER JOIN HethongTK ON SoDuKhachHang.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + TK.SoHieu + "*'"
    If TK.loai < 3 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKCNKH_ID, "0", CStr(TKCNKH_ID)) + " WHERE SoHieu LIKE '" + TK.SoHieu + "*'"
    If TK.loai > 2 Then ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKCNPT_ID, "0", CStr(TKCNPT_ID)) + " WHERE SoHieu LIKE '" + TK.SoHieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKVT()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 1
    shtk = FrmGetStr.GetString("S� hi�u TK", "��t/B� TK theo d�i chi ti�t")
    If Len(shtk) = 0 Then Exit Sub
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    If TK.tk_id = TKCNKH_ID Or TK.tk_id = TKCNPT_ID Or TK.tk_id = TSCD_ID Or TK.tk_id = KHTSCD_ID Or TK.tk_id = TKThue_ID Or TK.tk_id = TKDT_ID Then Exit Sub
    If TK.TkCoPS(0, 0) Or TK.NoDauKy <> 0 Or TK.CoDauKy <> 0 Then
        MsgBox "T�i kho�n c� ph�t sinh ho�c ��u k�, kh�ng chuy�n ��i ���c!", vbCritical, App.ProductName
        GoTo KT
    End If
    If TK.tk_id = TKVT_ID Then ExecuteSQL5 "DELETE TonKho.* FROM TonKho INNER JOIN HethongTK ON TonKho.MaTaiKhoan=HethongTK.MaSo WHERE HethongTK.SoHieu LIKE '" + TK.SoHieu + "*'"
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID=" + IIf(TK.tk_id = TKVT_ID, "0", CStr(TKVT_ID)) + " WHERE SoHieu LIKE '" + TK.SoHieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKDTTP()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 4
    shtk = FrmGetStr.GetString("S� hi�u TK", "��t/B� TK h�ch to�n doanh thu")
    If Len(shtk) = 0 Then GoTo KT
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Or Left(TK.SoHieu, 2) <> "51" Then GoTo KT
    If TK.TkCoPS(0, 0) Then
        MsgBox "T�i kho�n c� ph�t sinh, kh�ng chuy�n ��i ���c!", vbCritical, App.ProductName
        GoTo KT
    End If
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID2=" + IIf(TK.tk_id2 = TKDT_ID, "0", CStr(TKDT_ID)) + " WHERE SoHieu LIKE '" + TK.SoHieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub DatTKTS()
    Dim shtk As String, TK As New ClsTaikhoan
    
    FrmGetStr.tag = 3
    shtk = FrmGetStr.GetString("S� hi�u TK", "��t/B� TK theo d�i chi ti�t")
    If Len(shtk) = 0 Then Exit Sub
    TK.InitTaikhoanSohieu shtk
    If TK.MaSo = 0 Then GoTo KT
    ExecuteSQL5 "UPDATE HethongTK SET TK_ID2=" + IIf(TK.tk_id2 = TKCPSX_ID, "0", CStr(TKCPSX_ID)) + " WHERE SoHieu LIKE '" + TK.SoHieu + "*'"
KT:
    Set TK = Nothing
End Sub

Private Sub RunCT()
    Dim pctpath As String
    
    pctpath = SelectSQL("SELECT App1Path AS F1 FROM License")
    If Len(Dir(pctpath)) > 0 Then Shell pctpath, vbNormalFocus
End Sub

Public Function ChonTenTep(title As String, f As Long, mask As String, act As Integer) As String
    With dlgCommonDialog
        .DialogTitle = title
        .Flags = f
        .FileName = mask
        .DefaultExt = mask
        .Filter = "T�p d� li�u (" + mask + ")|" + mask + "|T�t c� (*.*)|*.*"
        On Error GoTo Xong
        Select Case act
            Case 1:            .ShowOpen
            Case 2:            .ShowSave
            Case 3:             .ShowPrinter
            Case 4:             .ShowFont
        End Select
        On Error GoTo 0
        If Len(.FileName) = 0 Or Left(.FileName, 1) = "*" Then GoTo Xong
        
        If act = 2 Then
            If Len(Dir(.FileName)) > 0 Then
                If .FileName = pDataPath Then
                    MsgBox "T�p d� li�u �ang m� !", vbCritical, App.ProductName
                    GoTo Xong
                End If
                If MsgBox("T�p " + .FileName + " �� t�n t�i, ti�p t�c ? !", vbQuestion + vbYesNo, App.ProductName) = vbNo Then GoTo Xong
                If Recycle(.FileName) <> 0 Then
                    MsgBox "Kh�ng xo� ���c t�p " + dlgCommonDialog.FileName + " !", vbExclamation, App.ProductName
                    GoTo Xong
                End If
            End If
        End If
        ChonTenTep = .FileName
    End With
Xong:
End Function

Private Function StationList() As Integer
    ' It is important that ReDim be used to define the array as the DLL,
' because the DLL depends on being able to redimension the array.
    ReDim msString(1) As String
' The array is 1-based rather than 0-based, regardless if Option Base 1
' is specified in the declarations section.
    Dim miLoop As Integer, i As Integer, LDBName As String, sql As String, U As String, X As String
 
    LDBName = Left(pDataPath, Len(pDataPath) - 3) + "LDB"
'    miLoop = LDBUser_GetUsers(msString, LDBName, OptLDBLoggedUsers)
' The function calls cannot be combined and must be used individually.
' Get the first user in the selected .LDB file.
    For i = 0 To miLoop - 1
        If i >= LBound(msString, 1) And i <= UBound(msString, 1) Then
            U = SelectSQL("SELECT TenNSD AS F1 FROM Users WHERE WS='" + msString(i) + "' AND TenNSD<>'" + X + "'")
            If U <> "0" Then
                sql = sql + Chr(13) + msString(i) + " : " + U
                X = U
            End If
        End If
    Next
    If miLoop > 1 Then
        lbCty(12).Caption = "C�c m�y tr�m: " + sql
    Else
        lbCty(12).Caption = ""
    End If
    
    StationList = miLoop
End Function

Private Sub LietKeTep()
    Dim i As Integer, fn As String, k As Integer
    
    For i = 1 To 5
        fn = GetSetting(IniPath, "RecentFiles", "File" + CStr(i))
        If Len(fn) > 0 And fn <> pDataPath Then
            'mnHT(17 + i).Caption = fn
           ' mnHT(17 + i).Visible = True
            k = k + 1
        Else
           ' mnHT(17 + i).Visible = False
        End If
    Next
    mnHT(23).Visible = (k > 0)
End Sub

Private Sub XoaQuery()
    Dim q As String
    
    q = InputBox("T�n query c�n xo�: ", App.ProductName)
    If Len(q) > 0 Then
        If QueryDaCo(q) Then DBKetoan.QueryDefs.Delete q
    End If
End Sub

Private Sub FontSetUp()
    Add32Font "VNTIME.TTF"
    Add32Font "VNTIMEB.TTF"
    Add32Font "VNTIMEBI.TTF"
    Add32Font "VNTIMEI.TTF"
    
    Add32Font "VHTIME.TTF"
    Add32Font "VHTIMEB.TTF"
    Add32Font "VHTIMEBI.TTF"
    Add32Font "VHTIMEI.TTF"
    
    Add32Font "VTIMESN.TTF"
    Add32Font "VTIMESB.TTF"
    Add32Font "VTIMESBI.TTF"
    Add32Font "VTIMESI.TTF"
End Sub


