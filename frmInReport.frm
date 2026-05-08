VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form frmInReport 
   BackColor       =   &H00FFFFC0&
   Caption         =   "In so va xuat file"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   BeginProperty Font 
      Name            =   "VK Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   4095
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFC0&
      Caption         =   " XuÊt file excel"
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   600
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  XuÊt file pdf"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Chän in"
      Height          =   255
      Left            =   3360
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Text            =   "Combo4"
      Top             =   600
      Width           =   3015
   End
   Begin VB.CheckBox chkSoCT 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sæ chi tiÕt tµi kho¶n"
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox chkSocai 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sæ c¸i tµi kho¶n"
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkSonk 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sæ nhËt ký chung"
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer timerImport 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   4200
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   8880
      TabIndex        =   9
      Text            =   "Combo3"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Chän tÊt c¶ tk"
      Height          =   195
      Left            =   5880
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2778
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "In kh«ng cÇn xem tr­íc"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8880
      TabIndex        =   5
      Text            =   "Combo3"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   3105
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3105
      Width           =   1095
   End
   Begin VB.CommandButton btnIn 
      Caption         =   "Thùc hiÖn"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3600
      Width           =   1230
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   7440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "DÕn th¸ng"
      Height          =   255
      Index           =   0
      Left            =   5160
      TabIndex        =   1
      Tag             =   "From"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tõ th¸ng"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   0
      Tag             =   "From"
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "frmInReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public itm As MSComctlLib.ListItem
Public typeprint As Integer
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Dim hWndApp As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
                                     ()

Private Sub GhiChutxt(ByVal content As Integer)
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
Private Sub Showform()
'Kiem tra ton tai file export.txt
    Dim FilePath As String
    Dim content As String
    FilePath = App.path & "\\HoaDon\\export.txt"
    Dim fileNumber As Integer

    Dim ch1 As Integer
    Dim ch2 As Integer
    Dim ch3 As Integer

    If chkSonk.Value = 1 Then
        ch1 = 1
    Else
        ch1 = 0
    End If
    If chkSocai.Value = 1 Then
        ch2 = 1
    Else
        ch2 = 0
    End If
    If chkSoCT.Value = 1 Then
        ch3 = 1
    Else
        ch3 = 0
    End If

    content = ch1 & "_" & ch2 & "_" & ch3 & "_" & Combo1.Text & "_" & Combo2.Text

    'Bo sung them check listview
    If chkSocai.Value = 1 Or chkSoCT.Value = 1 Then
        content = content & "_"
        For Each itm In ListView1.ListItems
            If itm.CHECKED Then
                content = content & "|" & itm.Text
            End If
        Next
    End If

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
    GhiChutxt 4
    Dim exePath As String
    exePath = App.path & "\\Tools\\Debug\\SaovietTax.exe"

    ' Shell d? m? ?ng d?ng
    Shell exePath, vbNormalFocus

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
End Sub
Private Sub CheckWindow()
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
Function CheckAndCreateFolder(ByVal folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Ki?m tra n?u thu m?c chua t?n t?i
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        CheckAndCreateFolder = True
        Debug.Print "Ðã t?o thu m?c: " & folderPath
    Else
        CheckAndCreateFolder = True
        Debug.Print "Thu m?c dã t?n t?i: " & folderPath
    End If
    
    Exit Function
    
ErrorHandler:
    CheckAndCreateFolder = False
    Debug.Print "L?i khi t?o thu m?c: " & Err.Description
End Function

Private Sub btnIn_Click()
'Tao thu muc Report truoc
'Lay ra nam chung tu
    Dim namct As Integer
    Dim reportpth As String
    namct = SelectSQL("SELECT NamTC AS F1 from License ")
    reportpth = App.path & "\\Tailieu\\Soketoan" & namct
    'Kiem tra folder ReportNamTC da co chua
    CheckAndCreateFolder (reportpth)

    'Lay ra danh sach tai khoan can in
    FBcKt.CboThang(0).Text = Combo1.Text
    FBcKt.CboThang(1).Text = Combo2.Text
    Dim prn As Printer
    Dim Port As String
    For Each prn In Printers
        If prn.DeviceName = cboPrinter.Text Then
            Port = prn.Port
        End If
    Next
    Dim preview As Integer

    If Check1.Value = 1 Then   ' ?? s?a l?i
        preview = 1   ' xem tru?c
    Else
        preview = 0  ' in
    End If
    typeprint = Combo3.ListIndex
    'Set Printer = typeprint   ' ?? QUAN TR?NG NH?T

    'Truong hop in excel
    If Option3.Value = True Then
        Showform
        Exit Sub
    End If

    ' Lay ra ten may in pdf
    Dim pdfPrinter As String
    For Each prn In Printers
        If InStr(1, prn.DeviceName, "PDF", vbTextCompare) > 0 Then
            pdfPrinter = prn.DeviceName
            Port = prn.Port
            Exit For    ' ?? l?y th?ng d?u tiên r?i thoát luôn
        End If
    Next
    Dim tenmayin As String
    'Truong hop co so nhat ky
    If Option1.Value = True Then
        tenmayin = Combo4.Text
    End If
    If Option2.Value = True Then
        tenmayin = pdfPrinter
    End If
    If chkSonk.Value = 1 Then
        'in bang may in hoac pdf
        If Option1.Value = True Or Option2.Value = True Then
            frmMain.typeprint = 2
            FBcKt.OptBC(100).Value = True
            FBcKt.Command_Click 0

            With frmMain.Rpt
                .PrinterDriver = "winspool"
                .printername = tenmayin
                .PrinterPort = Port  ' ?? port c?a b?n
                .Destination = preview
                .Action = 1
            End With
        End If
    End If
    'Cho so cai
    If chkSocai.Value = 1 Then
        For Each itm In ListView1.ListItems
            If itm.CHECKED Then
                'Debug.Print itm.Text   ' ?? tài kho?n dã ch?n
                frmMain.typeprint = 2
                FBcKt.OptBC(101).Value = True
                FBcKt.txtShTk(8).Text = itm.Text
                FBcKt.Command_Click 0
                With frmMain.Rpt
                    .PrinterDriver = "winspool"
                    .printername = tenmayin
                    .PrinterPort = Port  ' ?? port c?a b?n
                    .Destination = preview
                    .Action = 1
                End With
            End If
        Next
    End If
    'Cho so chi tiet tk
    If chkSoCT.Value = 1 Then
        For Each itm In ListView1.ListItems
            If itm.CHECKED Then

                'Debug.Print itm.Text   ' ?? tài kho?n dã ch?n
                frmMain.typeprint = 2
                FBcKt.Setinso (itm.Text)
                With frmMain.Rpt
                    .PrinterDriver = "winspool"
                    .printername = tenmayin
                    .PrinterPort = Port  ' ?? port c?a b?n
                    .Destination = preview
                    ' Thêm 2 dòng này d? in 2 m?t
                    '.PrinterDuplex = 2   ' 1=Simplex, 2=Horizontal, 3=Vertical
                    '.PrinterCopies = 1
                    .Action = 1
                End With
            End If
        Next
    End If
End Sub

Private Sub cboPrinter_Click()
    ExecuteSQL5 "UPDATE tbRegister SET printername = '" & cboPrinter.Text & "'"
End Sub

Private Sub Check2_Click()
    Dim dieukien As Boolean
    dieukien = Check2.Value
    For Each itm In ListView1.ListItems
        itm.CHECKED = dieukien
    Next
End Sub

Private Sub CreateDSMayin()
    Dim i As Integer
    Dim col As Integer
    Dim Row As Integer
    Dim Opt As OptionButton
    Dim prn As Printer

    Dim margin As Long
    Dim spacing As Long
    Dim itemWidth As Long

    margin = 200
    spacing = 200

    ' ?? width cho 2 item / hàng
    itemWidth = (Me.ScaleWidth - (margin * 2) - spacing) / 2

    i = 0
    col = 0
    Row = 0

    ' ===== DANH SÁCH PRINTER =====
    For Each prn In Printers

        ' ?? l?c printer rác
        If InStr(1, prn.DeviceName, "Fax", vbTextCompare) = 0 _
           And InStr(1, prn.DeviceName, "XPS", vbTextCompare) = 0 _
           And InStr(1, prn.DeviceName, "Microsoft Print to PDF", vbTextCompare) = 0 Then

            Set Opt = Me.Controls.Add("VB.OptionButton", "opt" & i, Me)

            With Opt
                .Caption = prn.DeviceName
                .BackColor = &HFFFFC0
                .Width = itemWidth
                .Left = margin + col * (itemWidth + spacing)
                .Top = 600 + Row * 400
                .Visible = True
            End With

            ' ?? c?p nh?t v? trí
            col = col + 1
            If col >= 2 Then
                col = 0
                Row = Row + 1
            End If

            i = i + 1
        End If

    Next

    ' ===== THÊM OPTION EXCEL =====
    Set Opt = Me.Controls.Add("VB.OptionButton", "opt" & i, Me)

    With Opt
        .Caption = " XuÊt file excel"
        .BackColor = &HFFFFC0
        .Width = itemWidth
        .Left = margin + col * (itemWidth + spacing)
        .Top = 550 + Row * 400
        .Visible = True
    End With

    col = col + 1
    If col >= 2 Then
        col = 0
        Row = Row + 1
    End If
    i = i + 1
End Sub
Private Sub Form_Load()
'Chi lay may in
    Dim prn As Printer
    For Each prn In Printers
        If InStr(1, prn.DeviceName, "PDF", vbTextCompare) = 0 _
           And InStr(1, prn.DeviceName, "XPS", vbTextCompare) = 0 _
           And InStr(1, prn.DeviceName, "Fax", vbTextCompare) = 0 Then
            Combo4.AddItem prn.DeviceName
            Combo4.Text = prn.DeviceName
        End If
    Next
    'CreateDSMayin
    Combo3.AddItem "Sæ nhËt ký chung"
    Combo3.AddItem "Sæ c¸i tµi kho¶n"
    Combo3.AddItem "Sæ chi tiÕt tµi kho¶n"
    Combo3.Text = "Sæ nhËt ký chung"

    With ListView1
        .View = lvwList        ' ?? d?i l?i cái này
        .CheckBoxes = True
    End With

    AddMonthToCbo Combo1
    AddMonthToCbo Combo2


    cboPrinter.Clear

    For Each prn In Printers
        cboPrinter.AddItem prn.DeviceName
    Next
    cboPrinter.AddItem "Excel"
    If cboPrinter.ListCount > 0 Then
        cboPrinter.Text = Printer.DeviceName
    End If
    'Lay gia tri may in tu database
    Dim printername As String
    printername = SelectSQL("SELECT Printername AS F1 from tbRegister ")
    If printername = "0" Or printername = "" Then
        ExecuteSQL5 "UPDATE tbRegister SET printername = '" & cboPrinter.Text & "'"
    Else
        cboPrinter.Text = printername
    End If

    ' ExecuteSQL5 "Update License set skiperror=0 where skiperror='...'"
    ' (Optional) ch?n máy in m?c d?nh
    'cboPrinter.Text = Printer.DeviceName

    Dim strSQL As String
    strSQL = "SELECT DISTINCTROW HeThongTK.SoHieu, HeThongTK.Cap, HeThongTK.Ten, HeThongTK.Kieu, HeThongTK.Loai, " & _
         "HeThongTK.DuNo_0 AS DkNo, HeThongTK.DuCo_0 AS DkCo, " & _
         "(HeThongTK.No_1 + HeThongTK.No_2 + HeThongTK.No_3 + HeThongTK.No_4 + " & _
         "HeThongTK.No_5 + HeThongTK.No_6 + HeThongTK.No_7 + HeThongTK.No_8 + " & _
         "HeThongTK.No_9 + HeThongTK.No_10 + HeThongTK.No_11 + HeThongTK.No_12) AS PsNo, " & _
         "(HeThongTK.Co_1 + HeThongTK.Co_2 + HeThongTK.Co_3 + HeThongTK.Co_4 + " & _
         "HeThongTK.Co_5 + HeThongTK.Co_6 + HeThongTK.Co_7 + HeThongTK.Co_8 + " & _
         "HeThongTK.Co_9 + HeThongTK.Co_10 + HeThongTK.Co_11 + HeThongTK.Co_12) AS PsCo, " & _
         "KC_N, KC_C, HeThongTK.DuNo_12 AS CkNo, HeThongTK.DuCo_12 AS CkCo " & _
         "FROM HeThongTK " & _
         "WHERE ((HeThongTK.MaTC = 0 Or HeThongTK.MaTC = HeThongTK.MaSo) OR (TK_ID3 Mod 10 >= 1)) " & _
         "And (HeThongTK.Loai > 0) AND Cap <= 1 " & _
         "AND (DuNo_12 <> 0 OR DuCo_12 <> 0 " & _
         "OR (HeThongTK.No_1 + HeThongTK.No_2 + HeThongTK.No_3 + HeThongTK.No_4 + " & _
         "HeThongTK.No_5 + HeThongTK.No_6 + HeThongTK.No_7 + HeThongTK.No_8 + " & _
         "HeThongTK.No_9 + HeThongTK.No_10 + HeThongTK.No_11 + HeThongTK.No_12) <> 0 " & _
         "OR (HeThongTK.Co_1 + HeThongTK.Co_2 + HeThongTK.Co_3 + HeThongTK.Co_4 + " & _
         "HeThongTK.Co_5 + HeThongTK.Co_6 + HeThongTK.Co_7 + HeThongTK.Co_8 + " & _
         "HeThongTK.Co_9 + HeThongTK.Co_10 + HeThongTK.Co_11 + HeThongTK.Co_12) <> 0) " & _
         "AND HeThongTK.Cap = 1 " & _
         "ORDER BY HeThongTK.SoHieu ASC"
    Dim rs_dstk As Recordset
    Set rs_dstk = DBKetoan.OpenRecordset(strSQL, dbOpenSnapshot)
    If Not rs_dstk.EOF Then
        Do While Not rs_dstk.EOF
            Set itm = ListView1.ListItems.Add(, , rs_dstk!sohieu)
            itm.CHECKED = True
            rs_dstk.MoveNext
        Loop
    End If
End Sub
