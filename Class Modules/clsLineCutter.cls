VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLineCutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsLineCutter

Private pLineLength As Double
Private pLayer As Layer
Private pLineThickness As Double
Private pLineColor As Long ' RGB dalam format Long
Private pAllShapes As Collection ' Koleksi untuk menyimpan shape dan garis yang diproses

' Properti untuk panjang garis
Public Property Let lineLength(value As Double)
    pLineLength = value
End Property

Public Property Get lineLength() As Double
    lineLength = pLineLength
End Property

' Properti untuk ketebalan garis
Public Property Let LineThickness(value As Double)
    pLineThickness = value
End Property

Public Property Get LineThickness() As Double
    LineThickness = pLineThickness
End Property

' Properti untuk warna garis
Public Property Let lineColor(value As Long)
    pLineColor = value
End Property

Public Property Get lineColor() As Long
    lineColor = pLineColor
End Property

' Properti untuk layer
Public Property Let Layer(value As Layer)
    Set pLayer = value
End Property

Public Property Get Layer() As Layer
    Set Layer = pLayer
End Property

' Method untuk menjalankan proses
Public Sub Run()
    Dim selectedShape As shape
    Set pAllShapes = New Collection

    ' Mengatur satuan pengukuran ke milimeter (mm)
    ActiveDocument.Unit = cdrMillimeter
    
    ' Pastikan ada shape yang dipilih
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih satu atau lebih shape terlebih dahulu!", vbExclamation
        Exit Sub
    End If

    ' Gunakan properti untuk panjang garis
    Dim lineLength As Double
    lineLength = Me.lineLength

    ' Buat layer khusus untuk line
    Set pLayer = ActivePage.CreateLayer("Line")
    
    ' Loop untuk setiap shape yang dipilih
    For Each selectedShape In ActiveSelection.Shapes
        Call AddLinesToShape(selectedShape)
    Next selectedShape

    ' Menyeleksi semua objek yang baru saja diproses
    Dim shape As shape
    For Each shape In pAllShapes
        shape.Selected = True
    Next shape

    ' Panggil method untuk menghapus duplikat dan menjaga satu garis per posisi
    RemoveAndKeepOne

End Sub

' Method untuk menambahkan garis pada setiap sudut shape
Private Sub AddLinesToShape(ByRef s As shape)
    Dim startX As Double, startY As Double
    Dim endX As Double, endY As Double
    Dim lineShape As shape
    Dim lineLength As Double: lineLength = Me.lineLength

    With s
        ' Sudut kiri atas - Garis horizontal
        startX = .leftX
        startY = .topY
        endX = startX + lineLength
        endY = startY
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape ' Menambahkan garis ke koleksi

        ' Sudut kiri atas - Garis vertikal
        startX = .leftX
        startY = .topY
        endX = startX
        endY = startY - lineLength
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kanan atas - Garis horizontal
        startX = .rightX - lineLength
        startY = .topY
        endX = .rightX
        endY = startY
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kanan atas - Garis vertikal
        startX = .rightX
        startY = .topY
        endX = startX
        endY = startY - lineLength
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kiri bawah - Garis horizontal
        startX = .leftX
        startY = .bottomY
        endX = startX + lineLength
        endY = startY
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kiri bawah - Garis vertikal
        startX = .leftX
        startY = .bottomY
        endX = startX
        endY = startY + lineLength
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kanan bawah - Garis horizontal
        startX = .rightX - lineLength
        startY = .bottomY
        endX = .rightX
        endY = startY
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape

        ' Sudut kanan bawah - Garis vertikal
        startX = .rightX
        startY = .bottomY
        endX = startX
        endY = startY + lineLength
        Set lineShape = pLayer.CreateLineSegment(startX, startY, endX, endY)
        ApplyLineProperties lineShape
        pAllShapes.Add lineShape
    End With
End Sub

' Method untuk menerapkan properti garis dengan menggunakan properti kelas
Private Sub ApplyLineProperties(ByRef lineShape As shape)
    With lineShape.Outline
        .width = Me.LineThickness ' Menggunakan ketebalan garis dari properti kelas

        ' Buat warna RGB dari Long
        Dim warnaRGB As New Color
        warnaRGB.RGBAssign _
            (Me.lineColor And &HFF), _
            ((Me.lineColor \ &H100) And &HFF), _
            ((Me.lineColor \ &H10000) And &HFF)

        .Color.CopyAssign warnaRGB
    End With
End Sub

' Method untuk menghapus duplikat
Private Sub RemoveAndKeepOne()
    Dim selShape As shape
    Dim shapesToDelete As Collection
    Set shapesToDelete = New Collection

    Dim objDict As Object
    Set objDict = CreateObject("Scripting.Dictionary")

    ' Memastikan ada objek yang diseleksi
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Tidak ada objek yang diseleksi!"
        Exit Sub
    End If

    ' Iterasi untuk setiap objek yang diseleksi
    For Each selShape In ActiveSelection.Shapes
        Dim key As String
        key = CStr(selShape.PositionX) & "-" & CStr(selShape.PositionY) & "-" & CStr(selShape.SizeWidth) & "-" & CStr(selShape.SizeHeight)

        ' Periksa apakah kunci sudah ada dalam dictionary
        If objDict.Exists(key) Then
            ' Tandai untuk dihapus jika sudah ada
            shapesToDelete.Add selShape
        Else
            objDict.Add key, selShape
        End If
    Next selShape

    ' Hapus semua objek yang sudah terdeteksi sebagai duplikat
    For Each selShape In shapesToDelete
        selShape.Delete
    Next selShape

End Sub







