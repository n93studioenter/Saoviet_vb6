VERSION 5.00
Begin VB.Form frmgioithieu 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Gioi thieu"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGioiThieu.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   5460
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Height          =   375
      Index           =   3
      Left            =   8280
      Picture         =   "frmGioiThieu.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "&Return"
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ph�n m�m ho�n to�n mi�n ph� cho doanh nghi�p nh�, c� s� ch�ng t� nh� h�n 500 v� doanh thu nh� h�n  05 t�."
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Xin vui l�ng li�n h�  090 575 7799 Mr. H�ng. "
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2520
      TabIndex        =   11
      Top             =   4680
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":6C04
      Height          =   615
      Index           =   6
      Left            =   360
      TabIndex        =   10
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Gi�i thi�u"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   9135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- L�m ���c nhi�u c�ng ty tr�n m�t ph�n m�m, th�m c�ng ty m�i ��n gi�n, ph� h�p cho ng��i d�ng l�m d�ch v� k� to�n."
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Chuy�n tr�c ti�p c�c b�o c�o ra m� v�ch."
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- M� c�c t�i kho�n chi ti�t ��n gi�n, ��ng k� theo d�i ��i t��ng nhanh."
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- Ph�n m�m ch� nh�p m�t l�n l� ra t�t c� c�c s�, t�ng h�p, chi ti�t, v�t t�, c�ng n�, c�ng tr�nh..."
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   10575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "- T�ng h�p, ng�n h�ng, v�t t�, b�n h�ng, TSC�, d�ng c�, k�t chuy�n t� ��ng..."
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ph�n m�m g�m c�c ph�n h�:"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   $"frmGioiThieu.frx":6CB7
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "frmgioithieu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command_Click(Index As Integer)
Unload Me
End Sub

