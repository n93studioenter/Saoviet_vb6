VERSION 5.00
Begin VB.Form Form4s 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form4"
   ScaleHeight     =   1305
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2760
      Top             =   2640
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form4s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Timer1.Interval = 2000
    Timer1.Enabled = True
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
    Dim tencty As String
    tencty = SelectSQL("select TenCty AS f1 from  License")

    Label1.Caption = "Da tai xong hoa don cho " & tencty
    Label1.FontName = "VNI-Times"
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
    Unload Me   ' ? Ðóng Form4
End Sub
