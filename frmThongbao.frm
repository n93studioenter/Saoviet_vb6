VERSION 5.00
Begin VB.Form frmThongbao 
   Caption         =   "Form4"
   ClientHeight    =   840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5235
   LinkTopic       =   "Form4"
   ScaleHeight     =   840
   ScaleWidth      =   5235
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "frmThongbao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
End Sub
Public Sub Thongbao(msg As String)
    Label1.Caption = msg
End Sub
