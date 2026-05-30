VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FThuChi 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Th«ng tin vÒ phiÕu thu - chi"
   ClientHeight    =   3540
   ClientLeft      =   3735
   ClientTop       =   3720
   ClientWidth     =   7680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FThuChi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Additional Voucher Information"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "0"
   Begin VB.CheckBox chkPhathanh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ph¸t hµnh"
      BeginProperty Font 
         Name            =   "VK Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   34
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Xem ho¸ ®¬n"
      Height          =   195
      Left            =   4080
      TabIndex        =   33
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   5
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   32
      Text            =   "..."
      Top             =   840
      Width           =   5415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   29
      Text            =   "Combo1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Timer timerImport 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   4920
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Ph¸t hµnh"
      Height          =   375
      Left            =   6000
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "L­u nh¸p"
      Height          =   375
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picFakeTitle 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   13575
      TabIndex        =   24
      Top             =   0
      Width           =   13575
      Begin VB.Label lblClose 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7200
         TabIndex        =   26
         Top             =   0
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   8430
         Index           =   0
         Left            =   720
         Picture         =   "FThuChi.frx":57E2
         Stretch         =   -1  'True
         Top             =   240
         Width           =   7890
      End
      Begin VB.Image picIcon 
         Appearance      =   0  'Flat
         Height          =   255
         Index           =   1
         Left            =   120
         Picture         =   "FThuChi.frx":112FF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "§¨ng nhËp"
         BeginProperty Font 
            Name            =   "VK Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   600
         TabIndex        =   25
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.CheckBox Checkinbangkevahoadon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In hãa ®¬n kÌm b¶ng kª"
      Height          =   255
      Left            =   3840
      TabIndex        =   23
      Tag             =   "Direct Export"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox checkinbangke 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In  b¶ng kª"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Tag             =   "Direct Export"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox CheckBox3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 3"
      Height          =   375
      Left            =   600
      TabIndex        =   21
      Top             =   4560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox CheckBox2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 2"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   4800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox CheckBox1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lien 1"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3240
      MaskColor       =   &H8000000A&
      TabIndex        =   19
      Top             =   5040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "In mÉu hoa ®¬n"
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Tag             =   "Direct Export"
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Néi dung theo diÔn gi¶i"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Tag             =   "Direct Export"
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   4
      Left            =   2040
      MaxLength       =   150
      TabIndex        =   13
      Text            =   "..."
      Top             =   120
      Width           =   5415
   End
   Begin VB.TextBox T 
      Height          =   345
      Index           =   3
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   4
      Text            =   "..."
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdkh 
      Height          =   375
      Left            =   3360
      Picture         =   "FThuChi.frx":115BC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtshkh 
      Height          =   345
      Left            =   2040
      LinkItem        =   "Sè hiÖu vËt t­ cÇn xem"
      MaxLength       =   20
      TabIndex        =   5
      Tag             =   "0"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   315
      Index           =   2
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   2
      Text            =   "..."
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   1
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   1
      Text            =   "..."
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox T 
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "..."
      Top             =   480
      Width           =   5415
   End
   Begin VB.CommandButton Command 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   6360
      Picture         =   "FThuChi.frx":11A36
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "&Save"
      Top             =   3000
      Width           =   1095
   End
   Begin MSMask.MaskEdBox MedNgay 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CCCD:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   31
      Tag             =   "Name of receiver,payer:"
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MÉu ho¸ ®¬n"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn C«ng ty:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Tag             =   "Name of receiver,payer:"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "H×nh thøc thanh to¸n"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Tag             =   "Object code"
      Top             =   2520
      Width           =   1575
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   330
      Left            =   2040
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2990;591"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "VNI-Times"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lbkh 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Tag             =   "1"
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hiÖu ®èi t­îng"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Tag             =   "Object code"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sè hãa ®¬n:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Tag             =   "Number of Voucher"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "§Þa chØ :"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Tag             =   "Address:"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tªn ng­êi nép tiÒn:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Tag             =   "Name of receiver,payer:"
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FThuChi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IdNhap As String
Dim countinvoiceinfo As Integer
Dim MaSoHd As Long
Dim tendotrangthai As Integer
Dim tendoidhoadon As String
Dim tendosotien As Double

' Khai báo c?u trúc STARTUPINFO và PROCESS_INFORMATION
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' Khai báo hàm CreateProcess, CloseHandle, WaitForSingleObject
Private Declare Function CreateProcess Lib "Kernel32" Alias "CreateProcessA" ( _
                                       ByVal lpApplicationName As String, _
                                       ByVal lpCommandLine As String, _
                                       ByVal lpProcessAttributes As Long, _
                                       ByVal lpThreadAttributes As Long, _
                                       ByVal bInheritHandles As Long, _
                                       ByVal dwCreationFlags As Long, _
                                       ByVal lpEnvironment As Long, _
                                       ByVal lpCurrentDirectory As String, _
                                       lpStartupInfo As STARTUPINFO, _
                                       lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetCurrentDirectory Lib "Kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = &HFFFF

Dim hWndApp As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                                    ()
Public FThuChiForm As Integer
Dim s(0 To 3) As String
Dim kh As New ClsKhachHang
Dim ngay As Date
Dim f1 As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
                                      ByVal hwnd As Long, _
                                      ByVal lpOperation As String, _
                                      ByVal lpFile As String, _
                                      ByVal lpParameters As String, _
                                      ByVal lpDirectory As String, _
                                      ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private typeGhichu As Integer

Public Sub RunExeViaWScript(ByVal exePath As String)
    Dim vbsPath As String
    Dim fnum As Integer

    ' T?o file VBS t?m
    vbsPath = Environ("TEMP") & "\run_" & Format(Timer, "0") & ".vbs"

    fnum = FreeFile
    Open vbsPath For Output As #fnum
    Print #fnum, "CreateObject(""WScript.Shell"").Run """"" & exePath & """"", 1, False"
    Close #fnum

    ' Ch?y VBS
    Shell "wscript.exe """ & vbsPath & """", vbHide

    ' Xóa file VBS sau 3 giây
    Shell "cmd.exe /c timeout /t 3 /nobreak > nul & del """ & vbsPath & """", vbHide
End Sub
Public Sub RunExeIndependent(ByVal exePath As String, Optional ByVal workingDir As String = "")
    If workingDir = "" Then
        workingDir = Left(exePath, InStrRev(exePath, "\"))
    End If
    
    ' Dùng "open" thay vì "runas" n?u không c?n admin
    ' "open" s? t?o process hoàn toàn d?c l?p, không k? th?a môi tru?ng t? VB6
    Dim ret As Long
    ret = ShellExecute(0, "open", exePath, "", workingDir, SW_NORMAL)
    
    If ret <= 32 Then
        MsgBox "L?i khi ch?y ?ng d?ng! Mã l?i: " & ret, vbCritical
    End If
End Sub

Public Sub Test()
 
    Unload Me
End Sub
Private Sub Timer1_Timer()
    Unload Me ' Ðóng form sau khi Timer h?t th?i gian
End Sub

Public Sub SaveGetPhieu()

'Lay dc maso hoa don MaSoHd
'Kiem tra co chua de insert or update
    Dim checkmaso As Integer
    Dim TyLe As String
    TyLe = "8%"
    checkmaso = SelectSQL("select count(*) AS f1 from tbGetphieu where MaCT = '" & MaSoHd & "'")
    If checkmaso = 0 Then
        'them moi
        ExecuteSQL5 "INSERT INTO tbGetphieu (MaCT, TenNM, TenCty, DiaChi, HTTT, CCCD, TyLe) VALUES (" & MaSoHd & ",'" & T(4).Text & "','" & T(0).Text & "','" & T(1).Text & "','" & ComboBox1.Text & "','" & T(5).Text & "','" & TyLe & "')"
    Else
        'Cap nhat
       ExecuteSQL5 "UPDATE tbGetphieu SET TenNM = '" & T(4).Text & "', TenCty = '" & T(0).Text & "', DiaChi = '" & T(1).Text & "', HTTT = '" & ComboBox1.Text & "', CCCD = '" & T(5).Text & "', TyLe = '" & TyLe & "' WHERE MaCT = '" & MaSoHd & "'"
    End If

End Sub

Public Sub Command_Click()
    Dim urlname As String
    urlname = SelectSQL("select Url AS f1 from  tbInvoiceInfo")
    If urlname = "seller-v2.tendoo.vn" Then
        If tendoidhoadon = "" Then
            Command1.Caption = "T¹o ®¬n hµng"
            chkPhathanh.Visible = True
            FrmChungtu.isInvoice = True
        End If
    End If
    'Cap nhat thong tin getphieu
    SaveGetPhieu

    If Combo1.ListIndex = -1 And Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If

    'Cho viettel
    If Combo1.ListCount > 0 And typeGhichu = 5 Then
        Dim id As Long
        id = Combo1.ItemData(Combo1.ListIndex)
        Dim FilePath As String
        FilePath = App.path & "\\HoaDon\\invoice.txt"
        Dim content As String
        content = FrmChungtu.txt(0).Text & "_" & FrmChungtu.txtVT(1).Text & "_" & FrmChungtu.MedNgay(0).Text & "_" & id
        Dim fileNumber As Integer
        If Not FileExists(FilePath) Then
            'Loai_thangbd_thangkt
            Dim iscreate As Boolean
            iscreate = CreateVersionFile(FilePath, content)
        Else
            fileNumber = FreeFile
            On Error Resume Next
            Open FilePath For Output As #fileNumber
            If Err.number = 0 Then
                Print #fileNumber, content;
                Close #fileNumber
                'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
            Else
                MsgBox "L?i khi ghi dè file!", vbExclamation
            End If
        End If
    End If

    'Cho bkav
    If typeGhichu = 6 Or typeGhichu = 8 Then
        FilePath = App.path & "\\HoaDon\\invoice.txt"
        content = FrmChungtu.txt(0).Text & "_" & FrmChungtu.txtVT(1).Text & "_" & FrmChungtu.MedNgay(0).Text

        If Not FileExists(FilePath) Then
            'Loai_thangbd_thangkt
            iscreate = CreateVersionFile(FilePath, content)
        Else
            fileNumber = FreeFile
            On Error Resume Next
            Open FilePath For Output As #fileNumber
            If Err.number = 0 Then
                Print #fileNumber, content;
                Close #fileNumber
                'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
            Else
                MsgBox "L?i khi ghi dè file!", vbExclamation
            End If
        End If
    End If


    Dim i As Integer
    '  ExecuteSQL5 "update chungtu set nguoimuahang = '" + T(4).Text + "'  where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    ExecuteSQL5 "update chungtu set nguoimuahang = '" + T(4).Text + "',hinhthucthanhtoan = '" + ComboBox1.Text + "',sophieudathang = '" + T(3).Text + "' ,chondiengiai = '" + str(Check1.Value) + "'  where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    For i = 0 To 3
        s(i) = T(i).Text
    Next
    FrmChungtu.Check1.Value = Check1.Value
    FrmChungtu.Check2.Value = Check2.Value
    FrmChungtu.hinhthucthanhtoan.Text = ComboBox1.Text + "  "
    FrmChungtu.thoihanthanhtoan.Text = MedNgay.Text
    FrmChungtu.sochungtu = T(3).Text
    FrmChungtu.CheckBox1 = CheckBox1.Value
    FrmChungtu.CheckBox2 = CheckBox2.Value
    FrmChungtu.CheckBox3 = CheckBox3.Value
    FrmChungtu.checkinbangke.Value = checkinbangke.Value
    FrmChungtu.Checkinbangkevahoadon.Value = Checkinbangkevahoadon.Value
    If FThuChi.FThuChiForm <> 0 Or countinvoiceinfo = 0 Or FrmChungtu.hdmoi = True Then
        Unload Me
    Else
        'Chi gianh cho viet tel, bkav

        Command1.Visible = True
        If IdNhap <> "" Then
            Command2.Visible = True
            Command2.Enabled = False
        End If
        'Check3.Visible = True
    End If
    If FThuChiForm = 1 Then
        ' FrmChungtu.DoneSetup
        FrmChungtu.timerNext.Enabled = True
    End If

    If FThuChiForm = 2 Then
        FrmChungtu.timerNext.Enabled = True
    End If

End Sub

Public Function FileExists(FilePath As String) As Boolean
    On Error GoTo ErrorHandler

    If Dir(FilePath) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If

    Exit Function

ErrorHandler:
    FileExists = False
End Function
Public Function CreateVersionFile(FilePath As String, content As String) As Boolean
    Dim fileNumber As Integer

    On Error GoTo ErrorHandler

    ' L?y file number
    fileNumber = FreeFile

    ' T?o file m?i (Output mode s? t?o file n?u chua có)
    Open FilePath For Output As #fileNumber

    ' Ghi n?i dung
    Print #fileNumber, content

    ' Ðóng file
    Close #fileNumber

    ' Ki?m tra file dã du?c t?o
    CreateVersionFile = FileExists(FilePath)

    Exit Function

ErrorHandler:
    CreateVersionFile = False
    On Error Resume Next
    Close #fileNumber
End Function
Public Sub GhiChutxt(ByVal content As Integer)
    Dim FilePath As String
    FilePath = App.path & "\\Hoadon\\status.txt"

    Dim FileNum As Integer
    FileNum = FreeFile  ' L?y s? file tr?ng

    Dim lineText As String
    Dim allText As String

    ' M? file d? d?c
    Open FilePath For Input As #FileNum

    ' Ð?c t?ng dòng d?n h?t file
    Do Until EOF(FileNum)
        Line Input #FileNum, lineText
        allText = allText & lineText & vbCrLf  ' N?i dòng và xu?ng dòng
    Loop

    ' Ðóng file
    Close #FileNum

    ' M? file d? ghi dè n?i dung
    FileNum = FreeFile    ' L?y l?i s? file tr?ng

    ' M? file d? ghi
    Open FilePath For Output As #FileNum
    Print #FileNum, content  ' Ghi n?i dung m?i (tham s? integer) vào file

    ' Ðóng file
    Close #FileNum
End Sub
Private Sub WriteMoreContent(ByVal FilePath As String, ByVal content As String)

    Dim fileNumber As Integer
    Dim oldContent As String
    Dim newContent As String

    On Error GoTo ErrorHandler

    '========================
    ' N?u file chua t?n t?i
    '========================
    If Dir(FilePath) = "" Then

        fileNumber = FreeFile

        Open FilePath For Output As #fileNumber

        Print #fileNumber, content;

        Close #fileNumber

        MsgBox "Ðã t?o file và ghi n?i dung!", vbInformation

    Else

        '========================
        ' Ð?c n?i dung cu
        '========================
        fileNumber = FreeFile

        Open FilePath For Input As #fileNumber

        oldContent = Input$(LOF(fileNumber), fileNumber)

        Close #fileNumber
        newContent = oldContent & "_" & content

        '========================
        ' Ghi thêm n?i dung m?i
        '========================
        fileNumber = FreeFile

        Open FilePath For Append As #fileNumber

        ' Xu?ng dòng r?i ghi thêm
        Print #fileNumber, vbCrLf & newContent;

        Close #fileNumber

        MsgBox "Ðã ghi thêm n?i dung thành công!", vbInformation

        ' Test xem n?i dung cu là gì
        Debug.Print oldContent

    End If

    Exit Sub

ErrorHandler:

    MsgBox "Có l?i x?y ra: " & Err.Description, vbExclamation

    On Error Resume Next

    If fileNumber <> 0 Then
        Close #fileNumber
    End If

End Sub
Private Sub Command1_Click()

    If chkPhathanh.Visible = True Then
        Dim fp As String
        Dim ct As String
        If chkPhathanh.Value = True Then
            ct = "1"
        Else
            ct = "0"
        End If
        fp = App.path & "\\HoaDon\\invoice.txt"
        Call WriteMoreContent(fp, ct)
    End If

    If tendotrangthai = True Then
        GhiChutxt typeGhichu
        Dim FilePath As String
        FilePath = App.path & "\\HoaDon\\invoice.txt"
        Dim content As String
        content = "TT_" & tendoidhoadon & "_" & tendosotien

        Dim fileNumber As Integer
        If Not FileExists(FilePath) Then
            'Loai_thangbd_thangkt
            Dim iscreate As Boolean
            iscreate = CreateVersionFile(FilePath, content)
        Else
            fileNumber = FreeFile
            On Error Resume Next
            Open FilePath For Output As #fileNumber
            If Err.number = 0 Then
                Print #fileNumber, content;
                Close #fileNumber
                'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
                DoEvents  ' Ð? d?m b?o ?ng d?ng có th?i gian kh?i d?ng
                Dim exePath2 As String
                Dim cmd2 As String

                exePath2 = App.path & "\Tools\Debug\SaovietTax.exe"
                ' Dùng runas v?i trust level th?p hon
                cmd2 = "runas /trustlevel:0x20000 """ & exePath2 & """"
                Shell cmd2, vbHide
            Else
                MsgBox "L?i khi ghi dè file!", vbExclamation
            End If
        End If
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass   ' ho?c 11
    Command_Click
    GhiChutxt typeGhichu
    Dim exePath As String
    Dim cmd As String

    exePath = App.path & "\Tools\Debug\SaovietTax.exe"
    ' Dùng runas v?i trust level th?p hon
    cmd = "runas /trustlevel:0x20000 """ & exePath & """"
    Shell cmd, vbHide
    Screen.MousePointer = vbDefault
    Exit Sub
    DoEvents  ' Ð? d?m b?o ?ng d?ng có th?i gian kh?i d?ng

    ' L?y handle c?a c?a s? ?ng d?ng
    hWndApp = 0  ' Kh?i t?o bi?n hWndApp

    While hWndApp = 0
        hWndApp = FindWindow(vbNullString, "frmMain")  ' Thay d?i tiêu d? c?a ?ng d?ng
        DoEvents  ' Cho phép x? lý s? ki?n khác
    Wend

    ' Ki?m tra handle có h?p l? hay không
    If hWndApp = 0 Then
        MsgBox "Không tìm th?y ?ng d?ng."
    Else
        ' Ð?i m?t chút tru?c khi ki?m tra l?i
        Sleep 1000
        CheckWindow
    End If
    Screen.MousePointer = vbDefault     ' ho?c 0

End Sub
Public Sub CheckWindow()
' Ki?m tra liên t?c xem c?a s? còn t?n t?i hay không
    Do
        If IsWindow(hWndApp) = 0 Then
            ' Ð?c file status.txt khi c?a s? không còn t?n t?i
            Dim FilePath As String
            FilePath = App.path & "\\Hoadon\\status.txt"

            Dim FileNum As Integer
            FileNum = FreeFile  ' L?y s? file tr?ng

            Dim lineText As String
            Dim allText As String

            ' M? file d? d?c
            Open FilePath For Input As #FileNum

            ' Ð?c t?ng dòng d?n h?t file
            Do Until EOF(FileNum)
                Line Input #FileNum, lineText
                allText = allText & lineText & vbCrLf  ' N?i dòng và xu?ng dòng
            Loop

            ' Ðóng file
            Close #FileNum

            ' Ki?m tra n?i dung file
            Dim textss As String
            textss = "ButtonClicked"
            Dim textss2 As String
            textss2 = SuperTrim(allText)

            If textss = textss2 Then
                timerImport.Enabled = True
            End If

            Exit Do
        End If
        DoEvents  ' Cho phép ?ng d?ng x? lý các s? ki?n khác
    Loop
End Sub
Function SuperTrim(ByVal s As String) As String
' Xóa t?t c? ký t? tr?ng (kho?ng tr?ng, tab, xu?ng dòng)
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCrLf, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    SuperTrim = Trim(s)  ' Xóa kho?ng tr?ng d?u/cu?i (ASCII 32)
End Function
Public Function RunExeClean(ByVal ExeFullPath As String, Optional ByVal workingDir As String = "", Optional ByVal WaitForExit As Boolean = False) As Boolean
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim lResult As Long

    ' 1. Thi?t l?p thu m?c làm vi?c (Working Directory)
    ' Vi?c này r?t quan tr?ng d? d?m b?o ?ng d?ng tìm dúng các file ph? tr?.

    ' N?u không truy?n vào thu m?c làm vi?c, t? d?ng l?y thu m?c c?a file exe
    If workingDir = "" Then
        workingDir = Left(ExeFullPath, InStrRev(ExeFullPath, "\"))
    End If
    ' Ð?t thu m?c làm vi?c cho ti?n trình hi?n t?i (VB6) tru?c khi t?o ti?n trình con
    Call SetCurrentDirectory(workingDir)

    ' 2. C?u hình STARTUPINFO
    si.cb = Len(si)
    ' wShowWindow = 1 (SW_SHOWNORMAL) d? hi?n th? c?a s? bình thu?ng
    si.wShowWindow = 1
    si.dwFlags = &H1    ' STARTF_USESHOWWINDOW

    ' 3. T?o ti?n trình m?i
    ' Luu ý: Tham s? lpCommandLine c?n ph?i n?m trong c?p d?u ngo?c kép n?u có d?u cách trong du?ng d?n
    lResult = CreateProcess(vbNullString, _
                            """" & ExeFullPath & """", _
                            0&, _
                            0&, _
                            1&, _
                            NORMAL_PRIORITY_CLASS, _
                            0&, _
                            workingDir, _
                            si, _
                            pi)

    If lResult = 0 Then
        RunExeClean = False
        Exit Function
    End If

    ' 4. (Tùy ch?n) Ch? cho ti?n trình k?t thúc
    If WaitForExit Then
        Call WaitForSingleObject(pi.hProcess, INFINITE)
    End If

    ' 5. D?n d?p Handle
    Call CloseHandle(pi.hProcess)
    Call CloseHandle(pi.hThread)

    RunExeClean = True
End Function
Private Sub Command2_Click()
    Screen.MousePointer = vbHourglass   ' ho?c 11
'Lay id phat hanh
    Dim rsports As Recordset

    Set rsports = DBKetoan.OpenRecordset("select IdNhap AS f1 FROM HoaDon " & _
                                         "inner join ChungTu on HoaDon.MaSo = ChungTu.MaSo " & _
                                         "where ChungTu.SoHieu = '" & FrmChungtu.txt(0).Text & "' " & _
                                         "and HoaDon.KyHieu = '" & FrmChungtu.txtVT(1).Text & "' " & _
                                         "and ChungTu.NgayCT = #" & Format(FrmChungtu.MedNgay(0).Text, "yyyy-mm-dd") & "#", dbOpenSnapshot)
    If Not rsports.EOF Then
        ' L?y giá tr? IdNhap
        Dim IdNhap As String
        If Not IsNull(rsports!f1) Then

            IdNhap = rsports!f1  ' ho?c rsport.Fields("f1").Value

            'Ghi noi dung file text invoice
            Dim FilePath As String
            FilePath = App.path & "\\HoaDon\\invoice.txt"
            Dim content As String
            content = "PH_" & IdNhap
            Dim fileNumber As Integer
            If Not FileExists(FilePath) Then
                'Loai_thangbd_thangkt
                Dim iscreate As Boolean
                iscreate = CreateVersionFile(FilePath, content)
            Else
                fileNumber = FreeFile
                On Error Resume Next
                Open FilePath For Output As #fileNumber
                If Err.number = 0 Then
                    Print #fileNumber, content;
                    Close #fileNumber
                    'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
                Else
                    MsgBox "L?i khi ghi dè file!", vbExclamation
                End If
            End If
            If typeGhichu = 5 Then
                GhiChutxt 5
            End If
            If typeGhichu = 6 Then
                GhiChutxt 6
            End If
            Dim exePath As String
            Dim cmd As String

            exePath = App.path & "\Tools\Debug\SaovietTax.exe"

            ' Dùng runas v?i trust level th?p hon
            cmd = "runas /trustlevel:0x20000 """ & exePath & """"

            Shell cmd, vbHide
            ' rsports.Close
            'Set rsports = Nothing
        Else

        End If

    End If
    Screen.MousePointer = vbDefault     ' ho?c 0

End Sub

Private Sub Form_Activate()
    If FThuChi.FThuChiForm = 5 Then
        Unload Me
    End If

    If Me.tag > 0 Then
        f1 = Me.tag
        Select Case f1
        Case 1:
            Label1(0).Caption = "Tªn ng­êi nhËn tiÒn:"
            Label1(1).Caption = "§Þa chØ ng­êi nhËn tiÒn:"
            T(2).Text = FrmChungtu.txt(0).Text
        Case 2:
            Me.Caption = "GiÊy Uû nhiÖm chi"
            Label1(0).Caption = "Tªn ®¬n vÞ nhËn tiÒn:"
            Label1(1).Caption = "Sè tµi kho¶n:"
            Label1(2).Caption = "T¹i Ng©n hµng:"
            T(0).MaxLength = 50
            T(2).MaxLength = 50
        Case 3:
            Me.Caption = "Ho¸ ®¬n b¸n hµng"
            Label1(0).Caption = "Tªn ng­êi mua hµng:"
            Label1(1).Caption = "§Þa chØ:"
            Label1(2).Caption = "H¹n thanh to¸n:"
            Label1(3).Caption = "Sè phiÕu ®Æt hµng:"
            T(2).Visible = False
            T(3).Visible = True
            txtshkh.Visible = False
            cmdkh.Visible = False
            lbkh.Visible = False
            MedNgay.Visible = True
        Case 10:
            Me.Caption = "Th«ng tin b¸o c¸o"
            Label1(0).Caption = "Ng­êi lËp biÓu:"
            Label1(1).Caption = "KÕ to¸n tr­ëng:"
            Label1(2).Caption = "Gi¸m ®èc:"
            Label1(3).Visible = False
            txtshkh.Visible = False
            cmdkh.Visible = False
            lbkh.Visible = False
            MedNgay.Visible = False
        End Select
        Me.tag = 0
    End If
    RFocus Command

    ComboBox1.AddItem ("Tieàn maët")
    ComboBox1.AddItem ("Chuyeån khoaûn")
    ComboBox1.AddItem ("Coâng nôï")
    ComboBox1.AddItem ("TM/CK")


    Dim sql As String
    Dim rs_chungtu As Recordset
    sql = "SELECT iif(Nguoimuahang is null ,'...',Nguoimuahang) as aa1,"
    sql = sql + "iif(hinhthucthanhtoan is null ,'...',hinhthucthanhtoan) as bb,"
    sql = sql + "iif(sophieudathang is null ,'...',sophieudathang) as cc,"
    sql = sql + "iif(chondiengiai is null ,'0',chondiengiai) as chon1  from chungtu where sohieu = '" + FrmChungtu.txt(0).Text + "'"
    Set rs_chungtu = DBKetoan.OpenRecordset(sql, dbOpenSnapshot)
    If rs_chungtu.recordCount > 0 Then
        T(4).Text = rs_chungtu!AA1
        T(3).Text = rs_chungtu!cc
        If rs_chungtu!chon1 = "2" Then
            Check1.Value = 1
        End If

        ComboBox1.Text = rs_chungtu!bb
    End If
    If FThuChiForm = 1 Or FThuChiForm = 3 Then
        Command_Click
    End If
    If FThuChiForm = 2 Then
        Command_Click
    End If
End Sub
Private Sub lblClose_Click()
    Unload Me
End Sub
Private Sub picFakeTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub lblTitle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFakeTitle_MouseDown Button, Shift, X, Y
End Sub
Private Sub Form_Load()
    Dim urlname As String
    urlname = SelectSQL("select Url AS f1 from  tbInvoiceInfo")
    If urlname = "seller-v2.tendoo.vn" Then

    End If
    Dim rsports As Recordset
    Set rsports = DBKetoan.OpenRecordset("select IdNhap AS f1, HoaDon.MaSo,HoaDon.TendoHDid,HoaDon.ThanhTien FROM HoaDon " & _
                                         "inner join ChungTu on HoaDon.MaSo = ChungTu.MaSo " & _
                                         "where ChungTu.SoHieu = '" & FrmChungtu.txt(0).Text & "' " & _
                                         "and HoaDon.KyHieu = '" & FrmChungtu.txtVT(1).Text & "' " & _
                                         "and ChungTu.NgayCT = #" & Format(FrmChungtu.MedNgay(0).Text, "yyyy-mm-dd") & "#", dbOpenSnapshot)
    If Not rsports.EOF Then
        ' L?y giá tr? IdNhap

        MaSoHd = rsports!MaSo
        If Not IsNull(rsports!TendoHDid) And rsports!TendoHDid <> "..." Then
            Command1.Caption = "CËp nhËt tr¹ng th¸i"
            Command1.Visible = True
            tendotrangthai = True
            tendoidhoadon = rsports!TendoHDid
            tendosotien = rsports!ThanhTien
        End If
        If Not IsNull(rsports!f1) And rsports!f1 <> "..." And rsports!f1 <> "" Then

            IdNhap = rsports!f1  ' ho?c rsport.Fields("f1").Value
            Command1.Visible = True
            Command2.Visible = True
            Command2.Enabled = True
            rsports.Close
            Set rsports = Nothing
        Else
        End If

    End If


    Dim countAccount As Integer
    countAccount = SelectSQL("select count(*) AS f1 from  tbInvoiceTemplate")
    If countAccount > 0 Then
        Combo1.Visible = True
        Label2.Visible = True
        Dim rstemplate As Recordset
        Set rstemplate = DBKetoan.OpenRecordset("SELECT DISTINCTROW tbInvoiceTemplate.* FROM tbInvoiceTemplate", dbOpenSnapshot)
        If Not rstemplate.EOF Then
            Combo1.Text = rstemplate!code & "-" & UnicodeToVni(rstemplate!Name)
        Else
            If Combo1.ListCount = 1 Then
                Combo1.Text = rstemplate!code & "-" & UnicodeToVni(rstemplate!Name)
            End If
        End If
        rstemplate.MoveNext
    End If
    countinvoiceinfo = SelectSQL("select count(*) AS f1 from  tbInvoiceInfo")
    'Neu chua co thi load danh sach
    If countinvoiceinfo > 0 Then
        Dim FilePath As String
        FilePath = App.path & "\\HoaDon\\invoice.txt"
        Dim content As String
        content = "1"
        Dim fileNumber As Integer
        If Not FileExists(FilePath) Then

            'Loai_thangbd_thangkt
            Dim iscreate As Boolean
            iscreate = CreateVersionFile(FilePath, content)
        Else
            fileNumber = FreeFile
            On Error Resume Next
            Open FilePath For Output As #fileNumber
            If Err.number = 0 Then
                Print #fileNumber, content;
                Close #fileNumber
                'MsgBox "Ðã ghi dè file version.txt thành công!", vbInformation
            Else
                MsgBox "L?i khi ghi dè file!", vbExclamation
            End If
        End If

        urlname = SelectSQL("select Url AS f1 from  tbInvoiceInfo")
        '5 la viettel
        Select Case urlname
        Case "vinvoice.viettel.vn"
            typeGhichu = 5
            GhiChutxt 5
        Case "van.ehoadon.vn"
            typeGhichu = 6
            GhiChutxt 6
        Case "id-v2.tendoo.vn"
            typeGhichu = 7
            GhiChutxt 7
        Case "seller-v2.tendoo.vn"
            typeGhichu = 8
            GhiChutxt 8
        Case Else
            ' Các tru?ng h?p còn l?i
            MsgBox "Url khác: " & urlname
        End Select
        If countAccount = 0 And typeGhichu = 5 Then
            Dim exePath As String
            exePath = App.path & "\\Tools\\Debug\\SaovietTax.exe"

            ' Shell d? m? ?ng d?ng
            Shell exePath, vbNormalFocus

            DoEvents  ' Ð? d?m b?o ?ng d?ng có th?i gian kh?i d?ng

            ' L?y handle c?a c?a s? ?ng d?ng
            hWndApp = 0  ' Kh?i t?o bi?n hWndApp
            Sleep 2000
            CheckWindow
        End If
    End If

    'Load danh sach cbb
    Dim rsport As Recordset
    Dim StartTime As Double
    Dim TimeoutSeconds As Integer
    Dim hasData As Boolean

    TimeoutSeconds = 10    ' Timeout sau 10 giây
    StartTime = Timer
    hasData = False

    If FThuChiForm = 0 And typeGhichu = 5 Then
        'Kiem tra xem co dang ky invoice nao khong
        Dim rsinvoice As Recordset
        Set rsinvoice = DBKetoan.OpenRecordset("SELECT * FROM tbInvoiceInfo", dbOpenSnapshot)
        If rsinvoice.EOF Then
            GoTo T
        Else
            If InStr(1, rsinvoice!url, "ehoadon", vbTextCompare) > 0 Then
                GoTo T
            End If
        End If
        ' L?p l?i query cho d?n khi có data ho?c timeout
        Set rsport = DBKetoan.OpenRecordset("SELECT DISTINCTROW tbInvoiceTemplate.* FROM tbInvoiceTemplate", dbOpenSnapshot)
        If Not rsport.EOF Then
            hasData = True
        End If
        Do While (Timer - StartTime) < TimeoutSeconds

            If Not rsport.EOF Then
                hasData = True
                Exit Do
            End If
            MsgBox rsport!id
            rsport.Close
            Set rsport = Nothing

            ' Ch? 0.5 giây tru?c khi query l?i
            Dim waitTime As Double
            waitTime = Timer
            Do While Timer - waitTime < 0.5
                DoEvents
            Loop
        Loop

        ' X? lý ComboBox
        If hasData Then
            ' Có data - add vào Combo nhu cu
            rsport.MoveFirst
            Do While Not rsport.EOF
                Combo1.AddItem rsport!code & "-" & UnicodeToVni(rsport!Name)
                Combo1.ItemData(Combo1.NewIndex) = rsport!id
                ' Ch? set text cho item d?u tiên
                'Kiem tra hoa don de lay template id
                Dim rshd As Recordset
                Set rshd = DBKetoan.OpenRecordset("SELECT  * FROM HoaDon where IdTemplate ='" & rsport!id & "' ", dbOpenSnapshot)
                If Not rshd.EOF Then
                    Combo1.Text = rsport!code & "-" & UnicodeToVni(rsport!Name)
                Else
                    If Combo1.ListCount = 1 Then
                        Combo1.Text = rsport!code & "-" & UnicodeToVni(rsport!Name)
                    End If
                End If
                rsport.MoveNext
            Loop

            ' Ðóng recordset
            rsport.Close
            Set rsport = Nothing
        Else
            ' Timeout - không có data
            MsgBox "Không có d? li?u sau " & TimeoutSeconds & " giây!", vbExclamation
        End If
    End If

T:
    lblTitle(11).AutoSize = True
    Me.Height = Me.Height + 350 + 10
    picFakeTitle.Width = Me.ScaleWidth
    picFakeTitle.Height = 325
    picIcon(1).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2
    lblTitle(11).Left = picIcon(1).Left + picIcon(1).Width + 90
    lblTitle(11).Top = (picFakeTitle.Height - picIcon(1).Height) \ 2 + 15
    lblClose.Top = 55
    lblTitle(11).Caption = "PhiÕu thu chi"
    AnControl Me
    Dim i As Integer

    For i = 0 To 3
        s(i) = "..."
    Next
    ngay = CVDate("01/01/1900")

    SetFont Me

End Sub

Private Sub T_GotFocus(Index As Integer)
    AutoSelect T(Index)
End Sub

Private Sub T_LostFocus(Index As Integer)
    If Len(T(Index).Text) = 0 Then T(Index).Text = "..."
End Sub

Public Sub GetPhieu(s1 As String, s2 As String, s3 As String, makh As Long, Optional d As Date, Optional s4 As String)
    kh.InitKhachHangMaSo makh
    'Lay danh sach get phieu
    Dim rsports As Recordset
    Set rsports = DBKetoan.OpenRecordset("select * from tbGetphieu where MaCT = '" & MaSoHd & "'", dbOpenSnapshot)
    If Not rsports.EOF Then
        T(5).Text = rsports!CCCD
        T(0).Text = rsports!tencty
        T(4).Text = rsports!tenNM
        T(1).Text = rsports!DiaChi
    Else
        T(0).Text = s1
        T(1).Text = s2
        T(2).Text = s3
        T(2).Text = FrmChungtu.txt(0).Text
        T(3).Text = s4
    End If

    txtshkh.Text = kh.sohieu
    lbkh.Caption = kh.Ten
    ngay = d
    If Year(d) > 1900 Then MedNgay.Text = Format(d, Mask_D)
    If Not Me.Visible Then
        Me.Show vbModal
    End If
    s1 = s(0)
    s2 = s(1)
    s3 = s(2)
    s4 = s(3)
    makh = kh.MaSo
    d = ngay
    Set kh = Nothing
    If FThuChiForm = 1 Then
        Command_Click
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbAltMask) > 0 Then
        Select Case KeyCode
            Case vbKeyG:
                RFocus Command
                Command_Click
        End Select
    End If
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtshkh_GotFocus()
    AutoSelect txtshkh
End Sub

Private Sub txtshkh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdkh_Click
End Sub

Private Sub txtshkh_LostFocus()
    Dim xxx As String, i As Integer
    
    kh.InitKhachHangSohieu txtshkh
    lbkh.Caption = kh.Ten
    If Len(T(0).Text) = 0 Or T(0).Text = "..." Then T(0).Text = kh.Ten
    Select Case f1
        Case 0, 1:
            If (Len(T(1).Text) = 0 Or T(1).Text = "...") And kh.DiaChi <> "..." Then T(1).Text = kh.DiaChi
        Case 2:
            xxx = Trim(LaySH(kh.taikhoan, 1, "-"))
            If (Len(T(1).Text) = 0 Or T(1).Text = "...") And IsNumeric(Left(xxx, 2)) Then T(1).Text = xxx
            i = Len(kh.taikhoan) - Len(xxx)
            If i > 0 Then
                xxx = Right(kh.taikhoan, i - 1)
                If (Len(T(2).Text) = 0 Or T(2).Text = "...") And Len(xxx) > 0 Then T(2).Text = xxx
            End If
    End Select
End Sub

Private Sub cmdkh_Click()
    Me.MousePointer = 11
    txtshkh.Text = FrmKhachHang.ChonKhachHang(txtshkh.Text)
    Me.MousePointer = 0
    RFocus txtshkh
End Sub

Private Sub MedNgay_GotFocus()
    AutoSelect MedNgay
End Sub

Private Sub MedNgay_LostFocus()
    If MedNgay.Text <> "__/__/__" Then
        If IsDate(MedNgay.Text) Then
            ngay = CDate(MedNgay.Text)
        Else
            RFocus MedNgay
        End If
    Else
        ngay = CVDate("01/01/1900")
    End If
End Sub

