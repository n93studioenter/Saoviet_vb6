VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTaiSan 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "T�i s�n"
   ClientHeight    =   7080
   ClientLeft      =   210
   ClientTop       =   510
   ClientWidth     =   9855
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Taisan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Fixed Assets Detail"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   19
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tr�ch kh�u hao t� th�ng ti�p theo"
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Tag             =   "Depreciate from month after increasing"
      Top             =   6120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   18
      Left            =   3840
      MaxLength       =   5
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "Taisan.frx":57E2
      Top             =   6060
      Width           =   495
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   14
      Left            =   5280
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   13
      Left            =   8205
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   12
      Left            =   6765
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   3840
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5330
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   5280
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   8205
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "Taisan.frx":57E6
      Top             =   4980
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   6765
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "Taisan.frx":57EA
      Top             =   4980
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   3840
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "Taisan.frx":57EE
      Top             =   4980
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   5280
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "Taisan.frx":57F2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   15
      Left            =   3840
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   6405
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   16
      Left            =   6765
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   6405
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   17
      Left            =   8205
      MaxLength       =   15
      MultiLine       =   -1  'True
      TabIndex        =   28
      Top             =   6405
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   8
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1920
      Width           =   6015
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   3960
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   5
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3000
      Width           =   6015
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   1
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   2
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   3
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   4
      Left            =   8520
      Picture         =   "Taisan.frx":57F6
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "&Print"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   8520
      Picture         =   "Taisan.frx":6C58
      Style           =   1  'Graphical
      TabIndex        =   38
      Tag             =   "&View"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   8520
      Picture         =   "Taisan.frx":7DCA
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "&Equipment"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   1
      Left            =   8520
      Picture         =   "Taisan.frx":810C
      Style           =   1  'Graphical
      TabIndex        =   36
      Tag             =   "&Return"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "Taisan.frx":952E
      Style           =   1  'Graphical
      TabIndex        =   35
      Tag             =   "&Save"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   0
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3480
      Width           =   3135
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   5
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3840
      Width           =   3495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   4
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3480
      Width           =   3495
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   6
      ItemData        =   "Taisan.frx":A95C
      Left            =   1440
      List            =   "Taisan.frx":A95E
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3840
      Width           =   3135
   End
   Begin VB.ComboBox Combo 
      Height          =   315
      Index           =   7
      ItemData        =   "Taisan.frx":A960
      Left            =   8640
      List            =   "Taisan.frx":A962
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   7320
      TabIndex        =   5
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      AutoTab         =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "99/99/99"
      PromptChar      =   "_"
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng�y CT"
      Height          =   255
      Index           =   34
      Left            =   6480
      TabIndex        =   69
      Tag             =   "Code"
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S� CT t�ng"
      Height          =   255
      Index           =   33
      Left            =   3960
      TabIndex        =   68
      Tag             =   "Code"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00E0E0E0&
      Caption         =   "n�m"
      Height          =   255
      Index           =   32
      Left            =   4440
      TabIndex        =   67
      Tag             =   "Year(s)"
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S� n�m t�nh KH"
      Height          =   255
      Index           =   31
      Left            =   240
      TabIndex        =   66
      Tag             =   "Year of Depreciation"
      Top             =   6090
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   11
      X1              =   9720
      X2              =   9720
      Y1              =   6960
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   120
      X2              =   120
      Y1              =   4320
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   9
      X1              =   120
      X2              =   9720
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   120
      X2              =   9720
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   28
      Left            =   6795
      TabIndex        =   65
      Top             =   5730
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   27
      Left            =   3720
      TabIndex        =   64
      Top             =   5760
      Width           =   1365
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   26
      Left            =   5280
      TabIndex        =   63
      Top             =   5760
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   22
      Left            =   2205
      TabIndex        =   62
      Top             =   6450
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   21
      Left            =   2205
      TabIndex        =   61
      Top             =   5730
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hao m�n trong th�ng :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   60
      Tag             =   "Monthly Depreciattion:"
      Top             =   6450
      Width           =   1620
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ngu�n kh�c"
      Height          =   255
      Index           =   11
      Left            =   6960
      TabIndex        =   59
      Tag             =   "Other"
      Top             =   4680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "T� b� sung"
      Height          =   255
      Index           =   10
      Left            =   3840
      TabIndex        =   58
      Tag             =   "Capital"
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ng�n s�ch"
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   57
      Tag             =   "Budget"
      Top             =   4680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gi� tr� c�n l�i :"
      Height          =   255
      Index           =   13
      Left            =   645
      TabIndex        =   56
      Tag             =   "Residual Value:"
      Top             =   5730
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Hao m�n :"
      Height          =   255
      Index           =   12
      Left            =   525
      TabIndex        =   55
      Tag             =   "Depreciation:"
      Top             =   5370
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nguy�n gi� :"
      Height          =   255
      Index           =   8
      Left            =   525
      TabIndex        =   54
      Tag             =   "Original Cost:"
      Top             =   5010
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "T�ng s�"
      Height          =   255
      Index           =   18
      Left            =   2685
      TabIndex        =   53
      Tag             =   "Total"
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   19
      Left            =   2205
      TabIndex        =   52
      Top             =   5010
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   20
      Left            =   2205
      TabIndex        =   51
      Top             =   5370
      Width           =   1455
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "T�n d�ng"
      Height          =   255
      Index           =   29
      Left            =   8445
      TabIndex        =   50
      Tag             =   "Credit"
      Top             =   4680
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,0"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   30
      Left            =   8160
      TabIndex        =   49
      Top             =   5730
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   7
      X1              =   8400
      X2              =   8400
      Y1              =   1440
      Y2              =   3360
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   5
      X1              =   120
      X2              =   8400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   120
      X2              =   8400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ghi ch� :"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   48
      Tag             =   "Notes"
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "N�ng l�c :"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   47
      Tag             =   "Ability"
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "T�n :"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   46
      Tag             =   "Description"
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S� hi�u :"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   45
      Tag             =   "Code"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "N�m s�n xu�t :"
      Height          =   255
      Index           =   24
      Left            =   240
      TabIndex        =   44
      Tag             =   "Pro. Year"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "N�m s� d�ng :"
      Height          =   255
      Index           =   25
      Left            =   2760
      TabIndex        =   43
      Tag             =   "Usage Year"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   3
      X1              =   8400
      X2              =   8400
      Y1              =   120
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   120
      X2              =   8400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   8400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "T�i kho�n :"
      Height          =   255
      Index           =   23
      Left            =   240
      TabIndex        =   42
      Tag             =   "Account"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lo�i :"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   41
      Tag             =   "Class"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nh�m :"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   40
      Tag             =   "Group"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "N��c s�n xu�t :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   34
      Tag             =   "Made in"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "T�nh tr�ng :"
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   33
      Tag             =   "State"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Qu�n l� :"
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   32
      Tag             =   "Managed by"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S� d�ng :"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   31
      Tag             =   "Used by"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Th�ng :"
      Height          =   255
      Index           =   17
      Left            =   8520
      TabIndex        =   30
      Tag             =   "Month"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Menu mnPU 
      Caption         =   "Danh �i�m"
      Visible         =   0   'False
      Begin VB.Menu mnDD 
         Caption         =   "N��c s�n xu�t..."
         Index           =   0
         Tag             =   "Country List..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "T�nh tr�ng s� d�ng..."
         Index           =   1
         Tag             =   "Conjucture List..."
      End
      Begin VB.Menu mnDD 
         Caption         =   "��i t��ng qu�n l�"
         Index           =   2
         Tag             =   "Administrative Object List"
      End
      Begin VB.Menu mnDD 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnDD 
         Caption         =   "H� th�ng t�i kho�n"
         Index           =   4
         Tag             =   "Chart of Account"
      End
   End
End
Attribute VB_Name = "frmTaiSan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TaiSan As New clsTaiSan

Dim KhoiTao As Integer
Dim pNhapdauky As Boolean
Dim psw As String
Dim ngay As Date
'======================================================================================
' FORM
'======================================================================================
' ACTIVATE : ��t d�ng tr�ng th�i v� thu�c t�nh MousePointer
Private Sub Form_Activate()
    pNhapdauky = (Me.tag > 0)
    If Not pNhapdauky Then
          SetListIndex Combo(7), Month(Date)
          DoEvents
          Combo(7).Enabled = True
          
          Chk.Visible = (pNghiepVu = NV_TANG)
          
          'MedNgay.Enabled = False
          'Text(19).Locked = True
    Else
          SetListIndex Combo(7), CLng(pThangDauKy)
          DoEvents
          Combo(7).Enabled = False
    End If
End Sub
' KEYDOWN : X� l� HotKey v� Escape
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      Dim i As Integer
      If (Shift And vbAltMask) > 0 Then
            i = -1
            Select Case KeyCode
                  Case vbKeyG: i = 0
                  Case vbKeyV:  i = 1
                  Case vbKeyP: i = 2
                  Case vbKeyX: i = 3
                  Case vbKeyI: i = 4
            End Select
            If i >= 0 Then
                If Command(i).Enabled Then
                        RFocus Command(i)
                        DoEvents
                        Command_Click (i)
                End If
            End If
      End If
      If KeyCode = vbKeyEscape Then Unload Me
End Sub

' LOAD
'     - L�y danh s�ch c�c ��i t��ng qu�n l�, s� d�ng, n��c s�n xu�t v� h� th�ng ph�n lo�i
'     - Kh�i t�o m�i tr��ng n�u nh�p m�i
'     - L�y v� hi�n th� n�i dung t�i s�n n�u �� c�
Private Sub Form_Load()
Dim chi_so As Integer
        InitDateVars MedNgay, ngay
      ' L�y danh s�ch c�c ��i t��ng quan h�
      Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM QuocGia ORDER BY Ten", Combo(0)
      Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan WHERE Cap = 1", Combo(1)
      Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM DTQly ORDER BY Ten", Combo(4)
      Int_RecsetToCbo "SELECT SoHieu + ' - ' + Ten AS F1, MaSo as F2 FROM HethongTK WHERE TK_ID2 = " + CStr(TKCPSX_ID) + " AND TKCon = 0 ORDER BY SoHieu", Combo(5)
      Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM TinhTrang ORDER BY Ten", Combo(6)
      AddMonthToCbo Combo(7)
      ' Kh�i t�o t�i s�n m�i
      If pMaTaiSan = 0 Then
            ' L�y danh s�ch c�c th�ng c� th� ch�n
            ' Kh�i t�o ��i t��ng TaiSan (Th� t�c n�y ph�i ���c g�i tr��c khi ��t th�ng ng�m ��nh)
            KhoiTaoTaiSan False
            ' N�u nh�p ��u k� th� ��t th�ng ng�m ��nh l� th�ng ��u k�, n�u t�ng th� ��t b�ng th�ng t�ng
            SetListIndex Combo(7), CLng(IIf(pNhapdauky, pThangDauKy, pThangTacDong))
            ' Kh�i t�o m�i tr��ng
            Command(2).Visible = False
            Command(3).Visible = False
            Command(4).Visible = False
            Label(12).Caption = "Hao m�n :"
            Label(16).Caption = "Kh�u hao / th�ng :"
      ' Hi�n th� c�c th�ng tin c�a t�i s�n �� c�
      Else
            ' L�y c�c th�ng tin trong d� li�u. Qu� tr�nh hi�n th� n�i dung t�i s�n c�n ph�i tr�nh c�c t�c ��ng
            ' do l�y v� ��t thu�c t�nh ListIndex cho c�c ComboBox l�m thay ��i thu�c t�nh ph�n lo�i �� c�
            KhoiTao = True
                  NoiDungTaiSan pMaTaiSan, pThangTacDong
            KhoiTao = False
            ' L�y danh s�ch c�c th�ng c� th� ch�n
            Do While chi_so < Combo(7).ListCount
                If Not InMonth(Combo(7).ItemData(chi_so), IIf(TaiSan.ThangTang = 0, pThangDauKy, TaiSan.ThangTang), IIf(TaiSan.ThangGiam = 13, IIf(pThangDauKy > 1, pThangDauKy - 1, 12), TaiSan.ThangGiam)) Then
                    Combo(7).RemoveItem chi_so
                Else
                    chi_so = chi_so + 1
                End If
            Loop
            ' ��t th�ng ng�m ��nh (s� d�n ��n vi�c hi�n th� c�c th�ng s� t��ng �ng)
            SetListIndex Combo(7), CLng(pThangTacDong)
            DoEvents
            ' Kh�i t�o m�i tr��ng
            Command(2).Visible = True
            Command(3).Visible = True
            Command(4).Visible = True
            Label(12).Caption = "T�ng hao m�n :"
            Label(16).Caption = "M�c kh�u hao th�ng :"
      End If
      
      If pMaTaiSan = 0 Then
            Me.Caption = " Nh�p t�i s�n m�i"
      Else
            Me.Caption = " S�a ��i chi ti�t t�i s�n"
      End If

      pGhichungtu = 0
      Caption = Caption + " - " + CStr(pNamTC)
      psw = GetSetting(IniPath, "Environment", "InvPsw")
      
      SetFont Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnPU
End Sub

' UNLOAD : Xo� tham chi�u ��n c�c Object �� kh�i t�o
Private Sub Form_Unload(Cancel As Integer)
      If pNhapdauky Then
            If KiemTraSoLieuDauKy = -1 Then
                  Cancel = True
                  Exit Sub
            End If
            pNhapdauky = False
      End If
      Set TaiSan.ThongSo = Nothing
      Set TaiSan = Nothing
End Sub
'======================================================================================
' command
'     1. Ghi t�i s�n
'           - Ki�m tra
'           - Ghi v�o d� li�u
'           - Chuy�n gi� tr� c�a t�i s�n v�o bi�n chung GiaTri �� ghi ch�ng t�
'           - Ghi c�c d�ng c� ph� t�ng k�m theo n�u c�
'           - N�u nh�p ��u k� th� t� ��ng t�o ch�ng t� ��u k�. N�u t�ng trong k� th� cho nh�p ch�ng t� t�ng
'           - N�u kh�ng ghi l�i ch�ng t� th� xo� t�i s�n v�a ghi.
'     2. In ho�c xem th� t�i s�n
'======================================================================================
Private Sub Command_Click(Index As Integer)
      Me.MousePointer = 11
      Select Case Index
            Case 0            ' Ghi t�i s�n ............................................................................................................................................
                  If TaiSan.HopLe = 0 Then
                        ' Th�m m�i (t�ng t�i s�n)
                        If pMaTaiSan = 0 Then
                              If TaiSan.ThemMoi(Chk.Value) = 0 Then
                                    pMaTaiSan = TaiSan.MaSo
                                    ' Chuy�n gi� tr� c�a t�i s�n v�a ghi v�o bi�n chung GiaTri �� t�o v� ghi ch�ng t�
                                    ' (ri�ng l��ng kh�u hao s� kh�ng ���c ghi v�o ch�ng t� t�ng ho�c ��u k� )
                                    With TaiSan.ThongSo
                                          GiaTri.NG_NS = .NG_NS
                                          GiaTri.NG_TBS = .NG_TBS
                                          GiaTri.NG_CNK = .NG_CNK
                                          GiaTri.NG_TD = .NG_TD
                                          GiaTri.CL_NS = .CL_NS
                                          GiaTri.CL_TBS = .CL_TBS
                                          GiaTri.CL_CNK = .CL_CNK
                                          GiaTri.CL_TD = .CL_TD
                                    End With
                                    ' Nh�p c�c d�ng c� ph� t�ng k�m theo t�i s�n
                                    Dim dem_
                                    dem_ = 0
                                    If vbYes = MsgBox("T�i s�n c� c�c d�ng c�, ph� t�ng k�m theo ?", _
                                                                                                                        vbYesNo + vbQuestion) Then
                                          frmDCPTung.Show 1
                                          frmTaiSan.Refresh
                                         Else
                                          dem_ = 1
                                    End If
                                    ' N�u nh�p ��u k� th� t�o ch�ng t� ��u k�, n�u t�ng trong k� th� t�o ch�ng t� t�ng
                                    If dem_ = 0 Then
                                          frmDCPTung.Show 1
                                          frmTaiSan.Refresh
                                    End If
                                    If pNhapdauky Then
                                          GhiChungTuDauKy
                                    Else
                                          ' Th�nh l�p d�ng ph�t sinh
                                          ThanhLapPhatSinh NV_TANG, TaiSan.MaTaiKhoan
                                          ' Ghi ch�ng t�
                                          pGhichungtu = 1
                                          SetListIndex FrmChungtu.CboThang, Combo(7).ItemData(Combo(7).ListIndex)
                                          FrmChungtu.txt(0).Text = Text(19).Text
                                          FrmChungtu.MedNgay(0).Text = MedNgay.Text
                                          FrmChungtu.MedNgay(1).Text = MedNgay.Text
                                          Unload Me
                                          Exit Sub
                                    End If
                                    pMaTaiSan = 0
                                    pMaChungTu = 0
                                    KhoiTaoTaiSan True
                              End If
                        ' S�a ��i (kh�ng c� nghi�p v� n�o k�m theo)
                        Else
                            If Len(psw) > 0 Then
                                If FPsw.GetPswX() <> psw Then GoTo XongTS
                            End If
                              Select Case TaiSan.SuaDoi
                                    Case 0
                                          pMaTaiSan = 0
                                          If Combo(7).ItemData(Combo(7).ListIndex) = pThangDauKy Then SoDuTKTS
                                          SendKeys "{Escape}", False
                                    Case -2, -3:
                                          MsgBox "Ch� � : s�a ��i l��ng kh�u hao v� gi� tr� c�a m�t t�i s�n �� b� gi�m ho�c ��nh " _
                                                       & "gi� l�i s� l�m cho s� li�u ghi tr�n ch�ng t� t��ng �ng kh�ng c�n ch�nh x�c n�a. " _
                                                       & "Xo� c�c ch�ng t� c� li�n quan �i v� sau �� ghi l�i n�u c�n s�a ��i", vbCritical
                              End Select
                        End If
                  End If
            Case 1            ' Tr� v� ...................................................................................................................................................
                  pMaTaiSan = 0
                  SendKeys "{Escape}", False
            Case 2            ' Danh s�ch d�ng c� ph� t�ng k�m theo .....................................................................................
                  frmDCPTung.Show 1
            Case 3            ' In th� t�i s�n trong n�m ...................................................................................................................
                  TaoTheTaiSan 0, Combo(7).ItemData(Combo(7).ListIndex)
            Case 4            ' Xem tr��c th� t�i s�n trong n�m ...................................................................................................
                  TaoTheTaiSan 1, Combo(7).ItemData(Combo(7).ListIndex)
      End Select
XongTS:
      Me.MousePointer = 0
End Sub
'======================================================================================
' COMBO
'======================================================================================
Private Sub Combo_Click(Index As Integer)
Dim i As Integer, vis As Boolean
      Select Case Index
            Case 0            ' N��c s�n xu�t
                  If Not Combo(0).ListIndex = -1 Then TaiSan.MaNuoc = Combo(0).ItemData(Combo(0).ListIndex) Else TaiSan.MaNuoc = 0
            Case 1            ' T�i kho�n
                  If Not Combo(1).ListIndex = -1 Then
                        If KhoiTao = False Then TaiSan.MaTaiKhoan = Combo(1).ItemData(Combo(1).ListIndex)
                        Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan WHERE CapTren = " + CStr(Combo(1).ItemData(Combo(1).ListIndex)), Combo(2)
                        If Combo(2).ListCount = 0 Then
                              TaiSan.maloai = 0
                              TaiSan.MaNhom = 0
                              Combo(3).Clear
                        End If
                   Else
                        TaiSan.MaTaiKhoan = 0
                  End If
            Case 2            ' Ph�n lo�i
                  If Not Combo(2).ListIndex = -1 Then
                        If KhoiTao = False Then TaiSan.maloai = Combo(2).ItemData(Combo(2).ListIndex)
                        Int_RecsetToCbo "SELECT SoHieu + '  ' + Ten AS F1, MaSo as F2 FROM LoaiTaiSan WHERE CapTren = " + CStr(Combo(2).ItemData(Combo(2).ListIndex)), Combo(3)
                        If Combo(3).ListCount = 0 Then
                              TaiSan.MaNhom = 0
                              TaoSoHieuTaiSan
                        End If
                   Else
                        TaiSan.maloai = 0
                  End If
            Case 3            ' Ph�n nh�m
                  If Not Combo(3).ListIndex = -1 Then
                        If KhoiTao = False Then TaiSan.MaNhom = Combo(3).ItemData(Combo(3).ListIndex)
                        TaoSoHieuTaiSan
                  Else
                        TaiSan.MaNhom = 0
                  End If
            Case 4            ' ��i t��ng qu�n l�
                  If Not Combo(4).ListIndex = -1 Then TaiSan.ThongSo.MaDTQL = Combo(4).ItemData(Combo(4).ListIndex) Else TaiSan.ThongSo.MaDTQL = 0
            Case 5            ' ��i t��ng s� d�ng
                  If Not Combo(5).ListIndex = -1 Then TaiSan.ThongSo.MaDTSD = Combo(5).ItemData(Combo(5).ListIndex) Else TaiSan.ThongSo.MaDTSD = 0
            Case 6            ' T�nh tr�ng s� d�ng
                  If Not Combo(6).ListIndex = -1 Then TaiSan.ThongSo.MaTTSD = Combo(6).ItemData(Combo(6).ListIndex) Else TaiSan.ThongSo.MaTTSD = 0
            Case 7            ' Th�ng
                  ' L�y th�ng t�ng c�a t�i s�n
                  If pMaTaiSan = 0 Then
                        If pNhapdauky Then
                              TaiSan.ThangTang = 0
                        Else
                              TaiSan.ThangTang = Combo(7).ItemData(Combo(7).ListIndex)
                              pThangTacDong = TaiSan.ThangTang
                        End If
                  ' Hi�n th� c�c th�ng s� t��ng �ng
                  Else
                        vis = (TaiSan.ThangTang = 0) And (pNghiepVu <> NV_TANG)
                        
                        MedNgay.Enabled = vis
                        Text(19).Locked = Not vis
                        MedNgay.TabStop = vis
                        Text(19).TabStop = vis
                        
                        TaiSan.ThongSo.ChiDinh pMaTaiSan, Combo(7).ItemData(Combo(7).ListIndex)
                        With TaiSan.ThongSo
                              ' Nguy�n gi�
                              Text(6).Text = Format(.NG_NS, Mask_0)
                              Text(7).Text = Format(.NG_TBS, Mask_0)
                              Text(8).Text = Format(.NG_CNK, Mask_0)
                              Text(9).Text = Format(.NG_TD, Mask_0)
                              ' Hao m�n
                              Text(10).Text = Format(.NG_NS - .CL_NS, Mask_0)
                              Text(11).Text = Format(.NG_TBS - .CL_TBS, Mask_0)
                              Text(12).Text = Format(.NG_CNK - .CL_CNK, Mask_0)
                              Text(13).Text = Format(.NG_TD - .CL_TD, Mask_0)
                              ' Kh�u hao
                              Text(14).Text = Format(.KH_NS, Mask_0)
                              Text(15).Text = Format(.KH_TBS, Mask_0)
                              Text(16).Text = Format(.KH_CNK, Mask_0)
                              Text(17).Text = Format(.KH_TD, Mask_0)
                              
                              'If (.KH_NS + .KH_CNK + .KH_TBS + .KH_TD) <> 0 Then
                              '      Text(18).Text = CStr(Fix((0.9 + (.NG_NS + .NG_CNK + .NG_TBS + .NG_TD) / (12 * (.KH_NS + .KH_CNK + .KH_TBS + .KH_TD)))))
                              'Else
                              '      Text(18).Text = "0"
                              'End If
                              
                              ' Gi� tr� c�n l�i
                              Label(26).Caption = Format(.CL_NS, Mask_0)
                              Label(27).Caption = Format(.CL_TBS, Mask_0)
                              Label(28).Caption = Format(.CL_CNK, Mask_0)
                              Label(30).Caption = Format(.CL_TD, Mask_0)
                              ' T�ng s�
                              Label(19).Caption = Format(.NG_NS + .NG_TBS + .NG_CNK + .NG_TD, Mask_0)
                              Label(20).Caption = Format((.NG_NS - .CL_NS) + (.NG_TBS - .CL_TBS) + (.NG_CNK - .CL_CNK) + (.NG_TD - .CL_TD), Mask_0)
                              Label(21).Caption = Format(.CL_NS + .CL_TBS + .CL_CNK + .CL_TD, Mask_0)
                              Label(22).Caption = Format(.KH_NS + .KH_TBS + .KH_CNK + .KH_TD, Mask_0)
                              SetListIndex Combo(4), .MaDTQL
                              SetListIndex Combo(5), .MaDTSD
                              SetListIndex Combo(6), .MaTTSD
                        End With
                        ' Cho ph�p s�a nguy�n gi� v� l��ng hao m�n c�a t�i s�n v�o th�ng ��u k�
                        ' (v�i �i�u ki�n t�i s�n ���c nh�p ��u k� v� ch�a b� ghi ch�ng t� gi�m)
'                        If (Combo(7).ListIndex = 0 Or Combo(7).ItemData(Combo(7).ListIndex) = TaiSan.ThangTang) And TaiSan.ThangGiam = 13 And TaiSan.ThangTang = 0 Then
                        If ((Combo(7).ListIndex = 0 Or Combo(7).ItemData(Combo(7).ListIndex) = TaiSan.ThangTang) And TaiSan.ThangGiam = 13 And Not KhongDC(TaiSan.MaSo)) Or (Combo(7).ListIndex = 0) Then
                              Text(6).Locked = False
                              Text(7).Locked = False
                              Text(8).Locked = False
                              Text(9).Locked = False
                              Text(10).Locked = False
                              Text(11).Locked = False
                              Text(12).Locked = False
                              Text(13).Locked = False
                        Else
                              Text(6).Locked = True
                              Text(7).Locked = True
                              Text(8).Locked = True
                              Text(9).Locked = True
                              Text(10).Locked = True
                              Text(11).Locked = True
                              Text(12).Locked = True
                              Text(13).Locked = True
                        End If
                  End If
                  ' Kh�ng cho s�a l��ng kh�u hao n�u t�i s�n �� gi�m trong n�m
                  If TaiSan.ThangGiam < 13 Then
                        Text(14).Locked = True
                        Text(15).Locked = True
                        Text(16).Locked = True
                        Text(17).Locked = True
                        Text(18).Locked = True
                  Else
                        Text(14).Locked = False
                        Text(15).Locked = False
                        Text(16).Locked = False
                        Text(17).Locked = False
                        Text(18).Locked = False
                  End If
      End Select
End Sub

Private Sub mnDD_Click(Index As Integer)
    Select Case Index
        Case 0:
            frmMain.mnTS_Click 3
            Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM QuocGia ORDER BY Ten", Combo(0)
        Case 1:
            frmMain.mnTS_Click 4
            Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM TinhTrang ORDER BY Ten", Combo(6)
        Case 2:
            frmMain.mnTS_Click 5
            Int_RecsetToCbo "SELECT Ten AS F1, MaSo as F2 FROM DTQly ORDER BY Ten", Combo(4)
        Case 4:
            FrmTaikhoan.tag = 1
            FrmTaikhoan.Show 1
            Int_RecsetToCbo "SELECT SoHieu + ' - ' + Ten AS F1, MaSo as F2 FROM HethongTK" _
                & " WHERE TK_ID2 = " + CStr(TKCPSX_ID) + " AND TKCon = 0 ORDER BY SoHieu", Combo(5)
    End Select
End Sub

'======================================================================================
' TEXT
'======================================================================================
' GotFocus
Private Sub Text_GotFocus(Index As Integer)
      AutoSelect Text(Index)
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 And (KeyAscii = 32 Or KeyAscii = 39 Or KeyAscii = 42) Then KeyAscii = 0
    If Index >= 6 And Index <= 18 Then KeyProcess Text(Index), KeyAscii
End Sub

' LostFocus
Private Sub Text_LostFocus(Index As Integer)
      Dim i As Integer, sn As Integer
      
      If Index >= 6 And Index <= 18 Then
            Text(Index).Text = Format(Text(Index).Text, Mask_0)
      End If
      If Len(Text(Index).Text) = 0 Then
            If (Index < 3 Or Index = 5) Then Text(Index).Text = "(...)" Else Text(Index).Text = "0"
      End If
      On Error GoTo Err_DataTypeConvertion
      Select Case Index
            Case 0: TaiSan.sohieu = Text(0).Text
            Case 1: TaiSan.Ten = Text(1).Text
            Case 2: TaiSan.NangLuc = Text(2).Text
            Case 3: TaiSan.NamSX = CInt5(Text(3).Text)
            Case 4: TaiSan.NamSD = CInt5(Text(4).Text)
            Case 5: TaiSan.GhiChu = Text(5).Text
            ' Nguy�n gi�
            Case 6: TaiSan.ThongSo.NG_NS = Cdbl5(Text(6).Text)
            Case 7: TaiSan.ThongSo.NG_TBS = Cdbl5(Text(7).Text)
            Case 8: TaiSan.ThongSo.NG_CNK = Cdbl5(Text(8).Text)
            Case 9: TaiSan.ThongSo.NG_TD = Cdbl5(Text(9).Text)
            ' Hao m�n
            Case 10: TaiSan.ThongSo.HM_NS = Cdbl5(Text(10).Text)
            Case 11: TaiSan.ThongSo.HM_TBS = Cdbl5(Text(11).Text)
            Case 12: TaiSan.ThongSo.HM_CNK = Cdbl5(Text(12).Text)
            Case 13: TaiSan.ThongSo.HM_TD = Cdbl5(Text(13).Text)
            ' Kh�u hao
            Case 14: TaiSan.ThongSo.KH_NS = Cdbl5(Text(14).Text)
            Case 15: TaiSan.ThongSo.KH_TBS = Cdbl5(Text(15).Text)
            Case 16: TaiSan.ThongSo.KH_CNK = Cdbl5(Text(16).Text)
            Case 17: TaiSan.ThongSo.KH_TD = Cdbl5(Text(17).Text)
            Case 18:
                                sn = CInt5(Text(18).Text)
                               TaiSan.NamKH = sn
                                If sn > 0 And (Not Text(18).Locked) Then
                                    For i = 6 To 9
                                         If Cdbl5(Label(20 + IIf(i < 9, i, 10)).Caption) > 0 Then
                                             Text(i + 8).Text = Format(RoundMoney(Cdbl5(Text(i).Text) / (sn * 12)), Mask_0)
                                             Text_LostFocus i + 8
                                         End If
                                    Next
                                End If
            Case 19:  TaiSan.shct = Text(19).Text
      End Select
      If Index > 5 And Index < 10 Then
            If Cdbl5(Text(Index).Text) = 0 Then
                Text(Index + 4).Text = "0"
                Text(Index + 8).Text = "0"
                Text_LostFocus Index + 4
                Text_LostFocus Index + 8
                Text(Index + 4).Enabled = False
                Text(Index + 8).Enabled = False
            Else
                Text(Index + 4).Enabled = True
                Text(Index + 8).Enabled = True
            End If
      End If
      On Error GoTo 0
      ' T�nh gi� tr� c�n l�i v� c�c t�ng s�
      With TaiSan.ThongSo
      If Index > 5 And Index < 14 Then
            If Index < 10 Then Label(19).Caption = Format(.NG_NS + .NG_TBS + .NG_CNK + .NG_TD, Mask_0) _
                                  Else Label(20).Caption = Format(.HM_NS + .HM_TBS + .HM_CNK + .HM_TD, Mask_0)
            .CL_NS = .NG_NS - .HM_NS
            .CL_TBS = .NG_TBS - .HM_TBS
            .CL_CNK = .NG_CNK - .HM_CNK
            .CL_TD = .NG_TD - .HM_TD
            Label(26).Caption = Format(.CL_NS, Mask_0)
            Label(27).Caption = Format(.CL_TBS, Mask_0)
            Label(28).Caption = Format(.CL_CNK, Mask_0)
            Label(30).Caption = Format(.CL_TD, Mask_0)
            Label(21).Caption = Format(.CL_NS + .CL_TBS + .CL_CNK + .CL_TD, Mask_0)
      End If
      If Index > 13 Then Label(22).Caption = Format(.KH_NS + .KH_TBS + .KH_CNK + .KH_TD, Mask_0)
      End With
      Exit Sub
Err_DataTypeConvertion:
      RFocus Text(Index)
End Sub
'======================================================================================
' SUB KhoiTaoTaiSan
'======================================================================================
Private Sub KhoiTaoTaiSan(tiep_tuc As Boolean)
Dim chi_so As Integer
      TaiSan.KhoiTao                                  ' Kh�i t�o ��i t��ng TaiSan
      For chi_so = 0 To 5                           ' Xo� c�c TextBox
            Text(chi_so).Text = ""
      Next
      For chi_so = 6 To 17                           ' Xo� c�c TextBox
            Text(chi_so).Text = "0"
      Next
      For chi_so = 19 To 30                         ' Xo� c�c Label
            If (chi_so >= 19 And chi_so <= 22) Or (chi_so >= 26 And chi_so <= 28) _
                                                                             Or chi_so = 30 Then Label(chi_so).Caption = "0"
      Next
      Combo(0).ListIndex = -1                     ' Xo� c�c Combo kh�ng thu�c h� th�ng ph�n lo�i
      For chi_so = 4 To 6
            Combo(chi_so).ListIndex = -1
      Next
      If tiep_tuc = False Then                      ' N�u l� l�n kh�i t�o ��u ti�n th� xo�
            For chi_so = 1 To 3                        ' c�c Combo thu�c h� th�ng ph�n lo�i
                  Combo(chi_so).ListIndex = -1
            Next
      Else                                                           ' N�u �ang ti�p t�c nh�p t�i s�n th� t�o s� hi�u m�i, l�y
            TaoSoHieuTaiSan                        ' m� s� c�a ph�n lo�i hi�n t�i v� th�ng t�ng ng�m ��nh
            TaiSan.MaTaiKhoan = Combo(1).ItemData(Combo(1).ListIndex)
            TaiSan.maloai = Combo(2).ItemData(Combo(2).ListIndex)
            If Combo(3).ListCount > 0 Then
                TaiSan.MaNhom = Combo(3).ItemData(Combo(3).ListIndex)
            Else
                TaiSan.MaNhom = 0
            End If
'            Combo_Click (7)
            RFocus Text(0)
      End If
      Combo_Click (7)
End Sub
'======================================================================================
' SUB NoiDungTaiSan : L�y v� hi�n th� n�i dung t�i s�n.
'                                 Ch� � : ��t thu�c t�nh ListIndex c�a c�c Combo thu�c h� th�ng ph�n lo�i s� d�n
'                                               ��n Events_Click t��ng �ng
'                                               N�i dung c�c th�ng s� s� ���c hi�n th� trong Events_Click c�a Combo(7)
'======================================================================================
Private Sub NoiDungTaiSan(ma_ts As Long, thang_cd As Integer)
Dim ml As Long, mn As Long
      TaiSan.ChiDinh ma_ts, thang_cd
      ml = TaiSan.maloai
      mn = TaiSan.MaNhom
      SetListIndex Combo(0), TaiSan.MaNuoc
      SetListIndex Combo(1), TaiSan.MaTaiKhoan
      SetListIndex Combo(2), ml
      SetListIndex Combo(3), mn
      TaiSan.maloai = ml
      TaiSan.MaNhom = mn
      Text(0).Text = TaiSan.sohieu
      Text(1).Text = TaiSan.Ten
      Text(2).Text = TaiSan.NangLuc
      Text(3).Text = TaiSan.NamSX
      Text(4).Text = TaiSan.NamSD
      Text(5).Text = TaiSan.GhiChu
      Text(18).Text = CStr(TaiSan.NamKH)
      Text(19).Text = TaiSan.shct
      ngay = TaiSan.NCT
      MedNgay.Text = Format(ngay, Mask_D)
End Sub
'========================================================================================================
' SUB TaoSoHieuTaiSan
'========================================================================================================
Private Sub TaoSoHieuTaiSan()
      Dim ms As Long, sql As String
      
      If Combo(3).ListCount > 0 Then
            ms = Combo(3).ItemData(Combo(3).ListIndex)
      Else
            ms = Combo(2).ItemData(Combo(2).ListIndex)
      End If
      sql = "SELECT SoHieu AS F1 FROM LoaiTaiSan WHERE MaSo = " + CStr(ms)
      Text(0).Text = CStr(SelectSQL(sql)) & "-"
End Sub
'======================================================================================
' SUB GhiChungTuDauKy : T�o ch�ng t� ri�ng cho c�c t�i s�n nh�p ��u k�.
'                                      Ch� � : Th�ng ch�ng t� ���c ghi b�ng 0, m� lo�i v� m� nh�m ���c ��t theo h�ng
'                                                     s� DK_LOAI v� DK_NHOM �� ph�n bi�t v�i c�c ch�ng t� kh�c.
'                                                     L��ng kh�u hao c�a ch�ng t� ��u k� lu�n b�ng 0.
'======================================================================================
Private Sub GhiChungTuDauKy()
Dim sql As String
      With GiaTri
      sql = "INSERT INTO CTTaiSan (MaSo,MaCTKT, SoHieu, Thang, VaoSo, NgayGhi, DienGiai, " _
            & "MaLoai, MaNhom, MaTS, NG_NS, NG_TBS, NG_CNK, NG_TD, " _
            & "CL_NS, CL_TBS, CL_CNK, CL_TD) VALUES (" + CStr(Lng_MaxValue("MaSo", "CTTaiSan") + 1) + ",0, '" + TaiSan.sohieu + "', 0" _
            + ",#" + Format(Date, Mask_DB) + "#,#" + Format(Date, Mask_DB) + "#,'" _
            + "Nh�p ��u k�" + "'," + CStr(DK_LOAI) + "," + CStr(DK_NHOM) + "," + CStr(pMaTaiSan) + "," _
            + DoiDau(.NG_NS) + "," + DoiDau(.NG_TBS) + "," + DoiDau(.NG_CNK) + "," + DoiDau(.NG_TD) + "," _
            + DoiDau(.CL_NS) + "," + DoiDau(.CL_TBS) + "," + DoiDau(.CL_CNK) + "," + DoiDau(.CL_TD) + ")" _
            + ""
      End With
      ExecuteSQL5 sql
End Sub
'======================================================================================
' SUB TaoTheTaiSan
'======================================================================================
Private Sub TaoTheTaiSan(ket_xuat As Integer, thang As Integer)
Dim rs_giam  As Recordset
Dim trong_nam As Double
Dim luy_ke As Double
Dim so_hieu As String
Dim ngay_thang As String
Dim dien_giai As String
Dim sql As String
      Me.MousePointer = 11
      HienThongBao " In th� t�i s�n", 1

      With GiaTri
            ' T�nh l��ng hao m�n cho ��n t�i th�ng hi�n t�i (c� tr�nh kh�u hao th�ng hi�n t�i)
            TinhGiaTriTaiSan TaiSan.MaSo, thang, KH_CO
            luy_ke = (.NG_NS - .CL_NS) + (.NG_TBS - .CL_TBS) + (.NG_CNK - .CL_CNK) + (.NG_TD - .CL_TD)
            ' T�nh l��ng hao m�n cho ��n h�t n�m (c� tr�nh kh�u hao th�ng cu�i n�m)
            If TaiSan.ThangTang > 0 Then
                trong_nam = luy_ke
            Else
                TinhGiaTriTaiSan TaiSan.MaSo, 0, KH_CO
                trong_nam = luy_ke - ((.NG_NS - .CL_NS) + (.NG_TBS - .CL_TBS) + (.NG_CNK - .CL_CNK) + (.NG_TD - .CL_TD))
            End If
      End With
      ' L�y ch�ng t� gi�m (n�u c�)
      sql = "SELECT SoHieu, NgayGhi, DienGiai FROM CTTaiSan WHERE MaTS = " _
                                                                  + CStr(TaiSan.MaSo) + " AND MaLoai = " + CStr(NV_GIAM)
      Set rs_giam = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      If rs_giam.EOF Then
            so_hieu = "....................."
            ngay_thang = "...................."
            dien_giai = ".............................................................."
      Else
            so_hieu = rs_giam!sohieu
            ngay_thang = Format(rs_giam!NgayGhi, "dd/mm/yy")
            dien_giai = rs_giam!diengiai
      End If
      rs_giam.Close
      Set rs_giam = Nothing
      ' D� li�u
      SetSQL "TheTaiSan", "SELECT DISTINCTROW TaiSan.SoHieu AS SoHieuTS, TaiSan.Ten AS TenTS, TaiSan.NangLuc, QuocGia.Ten AS TenNuoc, TaiSan.NamSX, TaiSan.NamSD, CTTaiSan.SoHieu AS SoHieuCT, CTTaiSan.Thang, CTTaiSan.DienGiai, (CTTaiSan.NG_NS+CTTaiSan.NG_TBS+CTTaiSan.NG_CNK+CTTaiSan.NG_TD) AS TNG " _
            & "FROM QuocGia RIGHT JOIN (TaiSan RIGHT JOIN CTTaiSan ON TaiSan.MaSo = CTTaiSan.MaTS) ON QuocGia.MaSo = TaiSan.MaNuoc " _
            & "WHERE TaiSan.MaSo = " + CStr(TaiSan.MaSo) + " ORDER BY CTTaiSan.NgayGhi"
      SetSQL "ThePhu", "SELECT DISTINCTROW DCPTung.Ten, DCPTung.DonVi, DCPTung.SoLuong, DCPTung.GiaThanh " _
            & "FROM DCPTung WHERE DCPTung.MaTS = " + CStr(TaiSan.MaSo) + " ORDER BY DCPTung.Ten"
      InTheTaiSan ket_xuat, "the1.rpt", Combo(4).List(Combo(4).ListIndex), _
                                                                        Format(trong_nam, Mask_0), Format(luy_ke, Mask_0)
      InTheTaiSan ket_xuat, "the2.rpt", so_hieu, ngay_thang, dien_giai
      Me.MousePointer = 0
End Sub
'======================================================================================
' SUB InTheTaiSan
'======================================================================================
Private Sub InTheTaiSan(ket_xuat As Integer, ten_file As String, _
                                                                                                                   ct_2 As String, ct_3 As String, ct_4 As String)
      ' K�t xu�t
      Select Case ket_xuat
            Case 0
                  frmMain.Rpt.Destination = 0
                  frmMain.Rpt.WindowTitle = "Th� t�i s�n c� ��nh"
            Case 1
                  frmMain.Rpt.Destination = 1
      End Select
      ' T�n File d� li�u v� b�o c�o
      SetRptInfo
      frmMain.Rpt.DataFiles(0) = pDataPath
      frmMain.Rpt.ReportFileName = ten_file
      ' C�ng th�c
      frmMain.Rpt.Formulas(0) = "TenCongTy = '" + pTenCty + "'"
      frmMain.Rpt.Formulas(1) = "TenChiNhanh = '" + pTenCn + "'"
      If ten_file = "the1.rpt" Then
            frmMain.Rpt.Formulas(2) = "QuanLy = '" + ct_2 + "'"
            frmMain.Rpt.Formulas(3) = "TrongNam = '" + ct_3 + "'"
            frmMain.Rpt.Formulas(4) = "LuyKe = '" + ct_4 + "'"
            frmMain.Rpt.Formulas(5) = "TenBaoCao = 'Th� t�i s�n c� ��nh'"
      Else
            frmMain.Rpt.Formulas(2) = "SoCT = '" + ct_2 + "'"
            frmMain.Rpt.Formulas(3) = "NgayThang = '" + ct_3 + "'"
            frmMain.Rpt.Formulas(4) = "LyDo = '" + ct_4 + "'"
            frmMain.Rpt.Formulas(5) = ""
      End If
      ' In b�o c�o
      InBaoCaoRPT
      Exit Sub
ErrorHandler:
      Beep
End Sub
'======================================================================================
' FUNCTION KiemTraSoLieuDauKy
'======================================================================================
Private Function KiemTraSoLieuDauKy() As Integer
Dim rs_dauky As Recordset, sql As String
      Me.MousePointer = 11
      sql = "SELECT Sum(NG_NS+NG_TBS+NG_CNK+NG_TD) AS TNG, " _
                              & "Sum(CL_NS+CL_TBS+CL_CNK+CL_TD) AS TCL " _
                              & "FROM CTTaiSan WHERE Thang=0"
      Set rs_dauky = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
      
      MsgBox "S� li�u ��u k� �� nh�p" + Chr(13) _
                  + "  T�ng nguy�n gi� : " + Format(rs_dauky!TNG, Mask_0) + Chr(13) _
                  + "  T�ng c�n l�i : " + Format(rs_dauky!TCL, Mask_0), vbInformation, App.ProductName
      KiemTraSoLieuDauKy = 0
      SoDuTKTS
      rs_dauky.Close
      Set rs_dauky = Nothing
      Me.MousePointer = 0
End Function

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub

Private Sub MedNgay_LostFocus()
    If IsDate(MedNgay.Text) Then
        ngay = CDate(MedNgay.Text)
        TaiSan.NCT = ngay
    Else
        RFocus MedNgay
    End If
End Sub
