VERSION 5.00
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "GRID32.OCX"
Begin VB.Form Formimport 
   Caption         =   "FormImport"
   ClientHeight    =   6645
   ClientLeft      =   75
   ClientTop       =   315
   ClientWidth     =   12705
   LinkTopic       =   "formImport"
   ScaleHeight     =   6645
   ScaleWidth      =   12705
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin MSGrid.Grid Grid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      _Version        =   65536
      _ExtentX        =   21616
      _ExtentY        =   7223
      _StockProps     =   77
      BackColor       =   16777215
   End
   Begin VB.CommandButton btnImport 
      Caption         =   "Xu ly"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
   End
End
Attribute VB_Name = "Formimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' FileInfo.cls
Option Explicit
Private fileList As Collection ' �? luu danh s�ch c�c d?i tu?ng FileInfo


Private Sub btnImport_Click()
    Dim FilePath As String
    Dim xmlDoc As Object
    Dim fDialog As Object
    Dim dlhDonNode As Object
    Dim ttChungNode As Object
    Dim ndhDonNode As Object
    Dim mstNode As Object
    Dim convertedDate As Date

    ' T?o h?p tho?i m? file
    Set fDialog = CreateObject("MSComDlg.CommonDialog")
    fDialog.ShowOpen
    FilePath = fDialog.fileName

    ' Kh?i t?o MSXML
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False

    ' T?i file XML
    If xmlDoc.Load(FilePath) Then
        ' L?y c�c node
        Set dlhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon")
        Set ttChungNode = xmlDoc.selectSingleNode("/HDon/DLHDon/TTChung")
        Set ndhDonNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon")
        Set mstNode = xmlDoc.selectSingleNode("/HDon/DLHDon/NDHDon/NMua/MST")

        ' Hi?n th? th�ng tin
        If Not dlhDonNode Is Nothing Then
            ' txtID.Text = dlhDonNode.Attributes.getNamedItem("Id").Text
        Else
            MsgBox "Kh�ng t�m th?y DLHDon."
        End If

        If Not ttChungNode Is Nothing Then
            Dim shDonNode As Object
            Dim shKHHDNode As Object
            Dim shNLapNode As Object

            Set shDonNode = ttChungNode.getElementsByTagName("SHDon")(0)
            Set shKHHDNode = ttChungNode.getElementsByTagName("KHHDon")(0)
            Set shNLapNode = ttChungNode.getElementsByTagName("NLap")(0)

            If Not shDonNode Is Nothing Then
                'txtSoHD.Text = shDonNode.Text
                'txtKHHD.Text = shKHHDNode.Text

                If Not shNLapNode Is Nothing Then
                    'convertedDate = CDate(shNLapNode.Text)
                    'txtDate.Text = Format(convertedDate, "dd/mm/yyyy")
                End If

                If Not mstNode Is Nothing Then
                    ' txtMST.Text = mstNode.Text
                    'GetcustomerByMST (txtMST.Text)
                Else
                    MsgBox "Kh�ng t�m th?y MST."
                End If
            Else
                MsgBox "Kh�ng t�m th?y SHDon."
            End If
        Else
            MsgBox "Kh�ng t�m th?y TTChung."
        End If

        If Not ndhDonNode Is Nothing Then
            ' X? l� ndhDonNode n?u c?n
        End If

    Else
        MsgBox "L?i khi t?i file XML: " & xmlDoc.parseError.reason
    End If
End Sub


Sub GetcustomerByMST(ByVal mst As String)
    Dim rs_ktra As Recordset
    Dim Query As String

    ' T?o truy v?n SQL d? l?y th�ng tin kh�ch h�ng theo MST
    Query = "SELECT Ten, DiaChi, MST FROM KhachHang WHERE MST = '" & mst & "'"

    ' M? Recordset d? l?y th�ng tin kh�ch h�ng
    Set rs_ktra = DBKetoan.OpenRecordset(Query, dbOpenSnapshot)

    ' Ki?m tra xem Recordset c� d? li?u kh�ng
    If Not rs_ktra.EOF Then
        ' Hi?n th? th�ng tin kh�ch h�ng
        MsgBox "T�n: " & rs_ktra.Fields("Ten").Value & vbCrLf & _
               "�?a Ch?: " & rs_ktra.Fields("DiaChi").Value & vbCrLf & _
               "MST: " & rs_ktra.Fields("MST").Value
    Else
        MsgBox "Kh�ng t�m th?y kh�ch h�ng v?i MST: " & mst
    End If

    ' ��ng Recordset
    rs_ktra.Close
    Set rs_ktra = Nothing
End Sub


Private Sub btnRefresh_Click()
 Set fileList = New Collection ' X�a v� kh?i t?o l?i danh s�ch
 TenThuTuc

End Sub

Private Sub Command1_Click()
 DuyetBangDuLieu
End Sub

Private Sub Form_Load()
    Set fileList = New Collection
    TenThuTuc
End Sub
Sub DuyetTatCaThuMuc(path As String)
    Dim folderStack As Collection
    Dim currentFolder As String
    Dim fileName As String
    Dim subFolderName As String
    Dim rowIndex As Long
    Dim gridHeaderAdded As Boolean ' Bi?n ki?m tra ti�u d? d� du?c th�m
    Dim fileDateCreated As String ' Bi?n d? luu ng�y t?o c?a file
       Dim lastFolderName As String ' Bi?n d? luu t�n thu m?c g?n nh?t ch?a file

    ' Kh?i t?o collection d? luu c�c thu m?c
    Set folderStack = New Collection
    
    ' �?m b?o du?ng d?n c� d?u "\" ? cu?i
    If Right(path, 1) <> "\" Then path = path & "\"
    
    ' Th�m thu m?c g?c v�o collection
    folderStack.Add path
    
    ' Kh?i t?o ch? s? h�ng cho Grid
    rowIndex = 0
    gridHeaderAdded = False ' Kh?i t?o l� False

    ' B?t d?u duy?t
    Do While folderStack.count > 0
        currentFolder = folderStack.item(1) ' L?y thu m?c d?u ti�n trong collection
        
        ' L?y t�n c�c file XML trong thu m?c hi?n t?i
        fileName = Dir(currentFolder & "*.xml")
        
        ' N?u chua th�m ti�u d?, th�m ti�u d? v�o khi t�m th?y file d?u ti�n
       
        ' Duy?t c�c file XML trong thu m?c hi?n t?i
        Do While fileName <> ""
            ' L?y ng�y t?o c?a file
            fileDateCreated = Format(FileDateTime(currentFolder & fileName), "dd/mm/yyyy") ' �?nh d?ng ng�y
             ' T�m t�n thu m?c g?n nh?t ch?a file
            
            Grid1.AddItem vbTab & fileName & vbTab & currentFolder & fileName & vbTab & fileDateCreated
            rowIndex = rowIndex + 1 ' Tang ch? s? h�ng
            
            fileName = Dir ' L?y file ti?p theo
        Loop
        
        ' L?y t�n thu m?c con c?a thu m?c hi?n t?i
        subFolderName = Dir(currentFolder & "*", vbDirectory)

        ' Duy?t qua c�c thu m?c trong thu m?c hi?n t?i
        Do While subFolderName <> ""
            If subFolderName <> "." And subFolderName <> ".." Then
                ' Ki?m tra xem d� c� ph?i thu m?c kh�ng
                If (GetAttr(currentFolder & subFolderName) And vbDirectory) = vbDirectory Then
                    ' Th�m thu m?c c?n duy?t v�o collection
                    folderStack.Add currentFolder & subFolderName & "\"
                End If
            End If
            subFolderName = Dir ' L?y thu m?c ti?p theo
        Loop

        ' X�a thu m?c d� duy?t kh?i collection
        folderStack.Remove 1
    Loop

    ' Th�ng b�o n?u kh�ng t�m th?y file n�o
    If rowIndex = 0 Then
        MsgBox "Kh�ng t�m th?y file XML n�o trong thu m?c d� ch? d?nh."
    End If
End Sub

Private Sub TenThuTuc()
    Dim folderPath As String
     Dim rowIndex As Integer

    ' �u?ng d?n d?n thu m?c g?c
    folderPath = "C:\TCP\S.T.E 24\S.T.E 24\Tailieu\Hoadonnchungtu\"

    ' Thi?t l?p ch? s? h�ng d?u ti�n trong Grid1
    rowIndex = 0
    ' G?i h�m d? quy d? duy?t qua thu m?c
    Grid1.Cols = 5
        Grid1.AddItem vbTab & "T�n File" & vbTab & "�u?ng D?n" & vbTab & "Ngay Tao"
        
        ' X�a t?t c? c�c m?c hi?n c� trong Grid
            Dim i As Long

    For i = Grid1.Rows - 1 To 0 Step -1
            If i > 2 Then
            Grid1.RemoveItem i ' X�a t?ng h�ng t? du?i l�n
             End If
        Next i
    DuyetTatCaThuMuc folderPath
End Sub
Sub DuyetBangDuLieu()
    Dim i As Long
    Dim j As Long
    Dim columnIndex As Long
    columnIndex = 0 ' Ch? s? c?t m� b?n mu?n duy?t (c� th? thay d?i n?u c?n)

    ' L?p qua t?t c? c�c h�ng trong Grid
     MsgBox Grid1.Text
End Sub
