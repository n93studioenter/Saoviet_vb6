VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form frmLocImport 
   Caption         =   "Import Data"
   ClientHeight    =   5520
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   12480
   LinkTopic       =   "Form4"
   ScaleHeight     =   5520
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12000
      Top             =   720
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Import"
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Text            =   "Combo2"
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lo�c"
      BeginProperty Font 
         Name            =   "VNI-Aptima"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin MSGrid.Grid Grid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   12015
      _Version        =   65536
      _ExtentX        =   21193
      _ExtentY        =   6588
      _StockProps     =   77
      BackColor       =   8454016
      Rows            =   20
      Cols            =   8
   End
   Begin VB.Label Label1 
      Caption         =   "�e�n"
      BeginProperty Font 
         Name            =   "VNI-Aptima"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "T��"
      BeginProperty Font 
         Name            =   "VNI-Aptima"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmLocImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public colSelect As String
Public rowSelect As String



Function FixVietnameseText(inputText As String) As String
' Danh s�ch c�c k� t? b? l?i v� k� t? d�ng tuong ?ng
    Dim charMap As Object
    Set charMap = CreateObject("Scripting.Dictionary")

    ' Th�m c�c c?p k� t? b? l?i v� k� t? d�ng
    ' Ki?m tra xem kh�a d� t?n t?i chua tru?c khi th�m
    If Not charMap.Exists("�") Then charMap.Add "�", "�"  ' Ch? � b? l?i th�nh �
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?
    If Not charMap.Exists("?") Then charMap.Add "?", "?"  ' Ch? ? b? l?i th�nh ?

    ' Duy?t qua t?ng k� t? trong chu?i d?u v�o v� thay th?
    Dim i As Long
    Dim outputText As String
    outputText = inputText

    For i = 1 To Len(inputText)
        Dim currentChar As String
        currentChar = Mid(inputText, i, 1)

        ' N?u k� t? hi?n t?i n?m trong danh s�ch b? l?i, thay th? b?ng k� t? d�ng
        If charMap.Exists(currentChar) Then
            outputText = Replace(outputText, currentChar, charMap(currentChar))
        End If
    Next i

    ' Tr? v? chu?i d� du?c s?a
    FixVietnameseText = outputText
End Function
Private Sub LocData(fromdate As Integer, todate As Integer)

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim folderPath As String
    Dim FilePath As String
    ' �u?ng d?n t?i thu m?c c?n l?y t?p
    folderPath = "C:\TCP\Saoviet\Hoadonchungtu"    ' Thay d?i du?ng d?n n�y theo thu m?c c?a b?n

    ' T?o d?i tu?ng FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ki?m tra xem thu m?c c� t?n t?i kh�ng
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)

        ' Duy?t qua t?ng c?t
        maxWidth = 2000
        'For i = 0 To Grid1.Cols - 1
        Grid1.ColWidth(0) = 1000
        Grid1.ColWidth(1) = 1000
        Grid1.ColWidth(2) = 3000
        Grid1.ColWidth(3) = 2000
        Grid1.ColWidth(4) = 2000
        Grid1.ColWidth(5) = 2000
        ' C?u h�nh MSFlexGrid
        With Grid1
            .Rows = 1    ' �?t l?i s? h�ng v? 1, ch? c�n l?i ti�u d? c?t
            .Cols = 7    ' S? c?t

            .AddItem "Ngay" & vbTab & "SoHD" & vbTab & "Ten Cty" & vbTab & "Dien giai" & vbTab & "Tong Tien" & vbTab & "No TK" & vbTab & "Co TK" & vbTab & "Ghi chu"      ' Th�m ti�u d? c?t
            'Clear List import
            FrmChungtu.ListReset
            ' Duy?t qua t?ng t?p trong thu m?c
            For Each file In folder.Files
                'Doc de lay ngay ra

                ' Kh?i t?o MSXML
                Dim xmlDoc As Object
                Dim ttChungNode As Object
                Dim shNLapNode As Object
                Dim TTNode As Object
                Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
                xmlDoc.async = False
                xmlDoc.validateOnParse = False
                FilePath = file.path

                If xmlDoc.Load(FilePath) Then
                    ' L?y c�c node
                    Dim shDonNode As Object
                    Dim shKHHDNode As Object
                    Dim ttNguoiBan As Object
                    Dim getMst As Object

                    Set ttNguoiBan = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/Ten")
                    Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")
                    Set getMst = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NBan/MST")
                    Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)
                    Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
                    Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)
                    Set TTNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/TToan/TgTCThue")
                    convertedDate = CDate(shNLapNode.Text)


 
                    ' Ki?m tra xem th�ng c?a convertedDate c� n?m trong kho?ng t? fromMonth d?n toMonth kh�ng
                    If Month(convertedDate) <= todate Then
                        'Them du lieu cho list frmChungtu
                        Dim getMaTKCo As String
                        Dim splitResult() As String
                        Dim tkno As String
                        Dim TkCo As String
                        Dim tkThue As String
                        Dim diengiai As String
                        tkno = ""
                        TkCo = ""
                        tkThue = ""
                        diengiai = ""

                        getMaTKCo = GetCusByMST(getMst.Text)
                        splitResult = Split(getMaTKCo, ",")
                        If getMaTKCo <> "" Then
                            tkno = splitResult(0)
                            TkCo = splitResult(1)
                            tkThue = splitResult(2)
                            diengiai = splitResult(3)
                        End If


                        FrmChungtu.AddImportData "", ttNguoiBan.Text, shDonNode.Text, Format(convertedDate, "dd/mm/yy"), "1", file.path, tkno, TkCo, tkThue, diengiai, "", "", "", "", "", "", "", ""
                        .AddItem Format(convertedDate, "dd/mm/yy") & vbTab & shDonNode.Text & vbTab & ttNguoiBan.Text & vbTab & diengiai & vbTab & Format(TTNode.Text, "#,##") & vbTab & tkno & vbTab & TkCo   ' Th�m d? li?u
                    End If
                End If
            Next file
        End With
    Else
        MsgBox "Thu m?c kh�ng t?n t?i!", vbExclamation
    End If

    ' Gi?i ph�ng b? nh?
    Set file = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Sub

Private Sub btnImport_Click()
    Me.Hide
    FrmChungtu.AutoCLickLoai
End Sub

Private Sub Command1_Click()
Dim fromdate As Integer
Dim todate As Integer
fromdate = Combo1.Text
todate = Combo2.Text
LocData fromdate, todate
End Sub
Function ChuanHoaChuoi(Text As String) As String
    ' Danh s�ch c�c k� t? b? l?i v� k� t? ti?ng Vi?t tuong ?ng
    Dim kyTuLoi() As String
    Dim kyTuChuan() As String
    Dim i As Integer
    
    ' Danh s�ch c�c k� t? b? l?i (lo?i b? c�c tru?ng h?p c? th?)
    kyTuLoi = Split("�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�", ",")
    
    ' Danh s�ch c�c k� t? ti?ng Vi?t tuong ?ng
    kyTuChuan = Split("�,?,?,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,�,a,d,i,u,o,u,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?", ",")
    
    ' Thay th? c�c k� t? b? l?i b?ng k� t? ti?ng Vi?t tuong ?ng
    For i = LBound(kyTuLoi) To UBound(kyTuLoi)
        Text = Replace(Text, CStr(kyTuLoi(i)), CStr(kyTuChuan(i)))
    Next i
    
    ' Tr? v? chu?i d� chu?n h�a
    ChuanHoaChuoi = Text
End Function
Function NormalizeString(ByVal InputString As String) As String
    ' Thay th? c�c k� t? kh�ng chu?n b?ng k� t? chu?n
    InputString = Replace(InputString, "�", "�")
    InputString = Replace(InputString, "?", "�") ' Thay th? d?u h?i th�nh �
    InputString = Replace(InputString, "�", "�")
    ' Th�m c�c thay th? c?n thi?t cho c�c k� t? kh�c
    InputString = Replace(InputString, "�", "O")
    InputString = Replace(InputString, "O", "O")
    InputString = Replace(InputString, "U", "U")
    
    ' Chuy?n v? ch? thu?ng d? so s�nh kh�ng ph�n bi?t ch? hoa/thu?ng
    NormalizeString = LCase(InputString)
End Function

Sub CompareStrings()
    Dim str1 As String
    Dim str2 As String
    
    str1 = "C�NG TY �I?N L?C B� R?A - VUNG T�U"
    str2 = "Co�ng Ty �ie�n L��c Ba� R�a - Vu�ng ta�u"

    ' Chu?n h�a chu?i
    str1 = NormalizeString(str1)
    str2 = NormalizeString(str2)

    ' So s�nh
    If str1 = str2 Then
        MsgBox "Hai chu?i gi?ng nhau."
    Else
        MsgBox "Hai chu?i kh�c nhau."
    End If
End Sub
Public Sub Test()
    Dim str1 As String
     str1 = "C�NG TY �I?N L?C B� R?A - VUNG T�U"
    str1 = NormalizeString(str1)
    Dim rs_ktra As Recordset
    Dim rs As DAO.Recordset
    Dim rst As String
    'Lay ra ma kh
    Query = "select * from KhachHang where MaSo  = 344"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? c�c b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            rst = rs_ktra!Ten
            Text2.Text = rst
            rs_ktra.MoveNext
        Loop
    End If
    '
    Dim chuoicoppy As String
    chuoicoppy = "C�NG TY �I?N L?C B� R?A - VUNG T�U"

    ' Khai b�o d?i tu?ng XML
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")

    ' C?u h�nh d?i tu?ng XML
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    ' Load file XML
    If xmlDoc.Load("C:\TCP\Saoviet\Hoadonchungtu\HD_PB15010034865_2025_2_1_1496315521_TD.xml") Then
        ' L?y ph?n t? <NBan> trong <NDHDon>
        Dim nBanNode As Object
        Set nBanNode = xmlDoc.selectSingleNode("//NDHDon/NBan")

        ' Ki?m tra xem ph?n t? <NBan> c� t?n t?i kh�ng
        If Not nBanNode Is Nothing Then
            ' L?y c�c gi� tr? t? c�c ph?n t? con c?a <NBan>
            Dim Ten As String
            Text1.Text = nBanNode.selectSingleNode("Ten").Text
            Ten = "sd"
            If Trim(Ten) = chuoicoppy Then
                MsgBox "trung nhau"
            End If
        End If
    Else
        MsgBox "Kh�ng th? load file XML. L?i: " & xmlDoc.parseError.reason
    End If
End Sub

Private Sub Form_Load()

    Dim i As Integer
    'cbb from
    Combo1.Clear
    ' V�ng l?p d? th�m th�ng t? 1 d?n 12
    For i = 1 To 12
        Combo1.AddItem i
    Next i
    Combo1.ListIndex = 0
    'cbb to
    Combo2.Clear
    ' V�ng l?p d? th�m th�ng t? 1 d?n 12
    For i = 1 To 12
        Combo2.AddItem i
    Next i
    Combo2.ListIndex = 11

    Command1_Click
End Sub

Function GetCusByMST(ByVal MaST As String) As String
    Dim numbers(1) As Integer    ' M?ng 2 ph?n t?

    Dim rs_ktra As Recordset
    Dim Query As String
    Dim rs As DAO.Recordset
    Dim fieldCount As Integer
    Dim i As Integer
    Dim rst As String

    'Lay ra ma kh
    Query = "select * from KhachHang where MST = '" & MaST & "'"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? c�c b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            rst = rs_ktra.Fields("MaSo").Value

            rs_ktra.MoveNext
        Loop
    Else
        rst = ""
    End If

    If rst = "" Then
        GetCusByMST = ""
    End If

    ' '''''''''''''''''''
    Query = "select * from HoaDon    where MaKhachHang  = " & CInt(rst) & " "
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? c�c b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            rst = rs_ktra.Fields("SoHD").Value
            ' Di chuy?n d?n b?n ghi ti?p theo
            rs_ktra.MoveNext
        Loop
    Else
        rst = ""
    End If

    If rst = "" Then
        GetCusByMST = ""
    End If
    ' ''''''''''''''''


    ' Lay MaTC tu bang chung tu
    Query = "SELECT TOP 2 MaTKNo,MaTKCo,Diengiai FROM ChungTu WHERE SoHieu =  '" & rst & "' ORDER BY MaSo DESC"
    'Query = "SELECT * from  ChungTu"
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)
    Dim Index As Integer
    Dim TkCo As Integer
    Dim tkno As Integer
    Dim tkThue As Integer
    Dim diengiai As String
    Index = 0
    If Not rs_ktra.EOF Then
        ' Duy?t qua t?t c? c�c b?n ghi
        Do While Not rs_ktra.EOF
            ' L?y s? lu?ng tru?ng
            If Index = 0 Then
                rst = rs_ktra.Fields("MaTKNo").Value
                tkThue = rst
            Else
                tkno = rs_ktra.Fields("MaTKNo").Value
                TkCo = rs_ktra.Fields("MaTKCo").Value
                diengiai = rs_ktra.Fields("Diengiai").Value
            End If

            ' Di chuy?n d?n b?n ghi ti?p theo
            ' Di chuy?n d?n b?n ghi ti?p theo
            Index = Index + 1
            rs_ktra.MoveNext
        Loop
    Else

    End If

    If rst <> "" Then

    Else
        GetCusByMST = rst
        Exit Function  ' Tho�t h�m

    End If

    ' '''''''''''''''''''''''''''''''''''''
    For i = 1 To 3

        ' T?o truy v?n SQL d? l?y th�ng tin kh�ch h�ng theo MST
        If i = 1 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & tkThue & " ORDER BY NgayKC DESC"
        End If
        If i = 2 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & tkno & " ORDER BY NgayKC DESC"
        End If
        If i = 3 Then
            Query = "SELECT TOP 1 * FROM HeThongTK WHERE MaTC = " & TkCo & " ORDER BY NgayKC DESC"
        End If


        'Query = "SELECT * from  ChungTu"

        ' M? Recordset d? l?y th�ng tin kh�ch h�ng
        Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

        If Not rs_ktra.EOF Then
            ' Duy?t qua t?t c? c�c b?n ghi
            Do While Not rs_ktra.EOF
                ' L?y s? lu?ng tru?ng
                rst = rs_ktra.Fields("SoHieu").Value
                If i = 1 Then
                    tkThue = rst
                End If
                If i = 2 Then
                    tkno = rst
                End If
                If i = 3 Then
                    TkCo = rst
                End If
                ' Di chuy?n d?n b?n ghi ti?p theo
                rs_ktra.MoveNext
            Loop
        Else
            GetCusByMST = rst
            Exit Function  ' Tho�t h�m
        End If

    Next i

    ' ��ng Recordset khi kh�ng c�n s? d?ng
    rs_ktra.Close
    Set rs_ktra = Nothing
    Dim result As String
    result = tkno & "," & TkCo & "," & tkThue & "," & diengiai
    
    GetCusByMST = result
End Function
Public Sub ChangeValueInpput(ByVal values As String)
    Grid1.Row = rowSelect
    Grid1.col = colSelect
    Grid1.Text = values
End Sub
Private Sub Grid1_DblClick()
' L?y gi� tr? � hi?n t?i
    Dim Value As String
    ' Gi? s? VBGrid1 l� t�n c?a Grid Control
    rowSelect = Grid1.Row
    colSelect = Grid1.col
    ftmInput.Text1.Text = Grid1.Text
    ftmInput.Show vbModal
    
    
    
    ' Hi?n th? gi� tr? �
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
