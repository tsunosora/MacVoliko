VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBannerFinishingLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsBannerFinishingLogic
Option Explicit

' -------------------------
' === Variabel untuk FitObject / Over ===
Private mFitSize As Double
Private mOverSize As Double
Private mUseOver As Boolean
Private mLineWidth As Double

' === Variabel untuk Tunnel ===
Private mTunnelSize As Double
Private mUseTunnelRL As Boolean

' Properti Publik untuk FitObject / Over
Public Property Let FitSize(val As Double)
    mFitSize = val
End Property

Public Property Let OverSize(val As Double)
    mOverSize = val
End Property

Public Property Let UseOver(val As Boolean)
    mUseOver = val
End Property

Public Property Let LineWidth(val As Double)
    mLineWidth = val
End Property

' Properti Publik untuk Tunnel
Public Property Let TunnelSize(val As Double)
    mTunnelSize = val
End Property

Public Property Let UseTunnelRL(val As Boolean)
    mUseTunnelRL = val
End Property

' -------------------------
' === Metode Apply Spacing (Untuk FitObject / Over)
Public Sub ApplySpacing()
    Dim doc As Document
    Dim s As shape
    Dim originalSelection As ShapeRange
    Dim extraSize As Double
    Dim w As Double, h As Double
    Dim posX As Double, posY As Double
    Dim rectangle As shape
    Dim w2 As Double, h2 As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double

    ' Tentukan ukuran tambahan (Fit atau Over)
    extraSize = IIf(mUseOver, mOverSize, mFitSize)
    
    If extraSize <= 0 Then
        MsgBox "Ukuran tambahan tidak valid!", vbExclamation
        Exit Sub
    End If

    ' Mendapatkan dokumen aktif
    Set doc = Application.ActiveDocument

    ' Memeriksa apakah ada shape yang dipilih
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih satu atau lebih shape terlebih dahulu!", vbExclamation
        Exit Sub
    End If

    ' Simpan seleksi awal
    Set originalSelection = ActiveSelectionRange

    doc.Unit = cdrCentimeter

    ' Loop untuk memproses setiap shape yang dipilih
    For Each s In ActiveSelection.Shapes
        w = s.SizeWidth
        h = s.SizeHeight
        posX = s.leftX
        posY = s.bottomY

        w2 = w + extraSize
        h2 = h + extraSize

        x1 = posX
        y1 = posY
        x2 = posX + w2
        y2 = posY + h2

        ' Membuat rectangle berdasarkan ukuran tambahan
        Set rectangle = doc.ActiveLayer.CreateRectangle(x1 - extraSize, y1 - extraSize, x2, y2)
        rectangle.Fill.ApplyNoFill
        rectangle.Outline.width = mLineWidth / 10
        rectangle.Outline.Color.RGBAssign 0, 0, 0
    Next s

    ' Seleksi ulang shape yang dipilih di awal
    doc.ClearSelection
    originalSelection.CreateSelection

End Sub

' -------------------------
' === Metode Apply Tunnel (Untuk Tunnel dengan pilihan atas/bawah atau kanan/kiri)
Public Sub ApplyBasedOnOption()
    ' Menentukan apakah menggunakan Tunnel RL atau TB
    If mUseTunnelRL Then
        Call ApplyTunnelRL
    Else
        Call ApplyTunnelTB
    End If
End Sub

' -------------------------
' === Metode Apply Tunnel - Atas/Bawah (Vertical Tunnel)
Private Sub ApplyTunnelTB()
    Dim doc As Document
    Dim s As shape
    Dim originalSelection As ShapeRange
    Dim w As Double, h As Double
    Dim posX As Double, posY As Double
    Dim rectangle1 As shape, rectangle2 As shape, rectangle3 As shape
    Dim L As Double, K As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim Ky1 As Double, Ky2 As Double

    ' Mendapatkan dokumen aktif
    Set doc = Application.ActiveDocument
    
    ' Memeriksa apakah ada shape yang dipilih
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih shape!", vbExclamation
        Exit Sub
    End If
    ' Simpan seleksi awal
    Set originalSelection = ActiveSelectionRange

    doc.Unit = cdrCentimeter

    ' Memastikan ukuran Tunnel valid
    If mTunnelSize <= 0 Then
        MsgBox "Ukuran Tunnel tidak valid!", vbExclamation
        Exit Sub
    End If

    L = 2
    K = mTunnelSize + L

    ' Loop untuk setiap shape yang dipilih
    For Each s In ActiveSelection.Shapes
        w = s.SizeWidth
        h = s.SizeHeight
        posX = s.leftX
        posY = s.bottomY

        x1 = posX
        y1 = posY
        x2 = posX + w
        y2 = posY + h
        Ky1 = y1 - K
        Ky2 = y2 + K

        ' Membuat tiga rectangle untuk tunnel
        Set rectangle1 = doc.ActiveLayer.CreateRectangle(x1, Ky1, x2, Ky2)
        Set rectangle2 = doc.ActiveLayer.CreateRectangle(x1, Ky2 - L, x2, Ky2)
        Set rectangle3 = doc.ActiveLayer.CreateRectangle(x1, Ky1, x2, Ky1 + L)

        ApplyStyle rectangle1
        ApplyStyle rectangle2
        ApplyStyle rectangle3
    Next s
        ' Seleksi ulang shape yang dipilih di awal
        doc.ClearSelection
        originalSelection.CreateSelection
End Sub

' -------------------------
' === Metode Apply Tunnel - Kiri/Kanan (Horizontal Tunnel)
Private Sub ApplyTunnelRL()
    Dim doc As Document
    Dim s As shape
    Dim originalSelection As ShapeRange
    Dim posX As Double, posY As Double
    Dim w As Double, h As Double
    Dim r As Double, a As Double, a2 As Double, w2 As Double
    Dim rectangle1 As shape, rectangle2 As shape, rectangle3 As shape
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double

    ' Mendapatkan dokumen aktif
    Set doc = Application.ActiveDocument
    
    ' Memeriksa apakah ada shape yang dipilih
    If ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Pilih shape!", vbExclamation
        Exit Sub
    End If
    ' Simpan seleksi awal
    Set originalSelection = ActiveSelectionRange

    doc.Unit = cdrCentimeter

    ' Memastikan ukuran Tunnel valid
    If mTunnelSize <= 0 Then
        MsgBox "Ukuran Tunnel tidak valid!", vbExclamation
        Exit Sub
    End If

    r = mTunnelSize + 2

    ' Loop untuk setiap shape yang dipilih
    For Each s In ActiveSelection.Shapes
        w = s.SizeWidth
        h = s.SizeHeight
        posX = s.leftX
        posY = s.bottomY

        x1 = posX
        y1 = posY + h
        x2 = posX + w
        y2 = posY

        a = r - 2
        a2 = r - a
        w2 = w + (r - a2)

        ' Membuat tiga rectangle untuk tunnel
        Set rectangle1 = doc.ActiveLayer.CreateRectangle(x1 - r, y1, x2 + r, y2)
        Set rectangle2 = doc.ActiveLayer.CreateRectangle(x1 - r, y1, x2 - w2, y2)
        Set rectangle3 = doc.ActiveLayer.CreateRectangle(x1 + w2, y1, x2 + r, y2)

        ApplyStyle rectangle1
        ApplyStyle rectangle2
        ApplyStyle rectangle3
    Next s
        ' Seleksi ulang shape yang dipilih di awal
        doc.ClearSelection
    originalSelection.CreateSelection
End Sub

' -------------------------
' === Metode untuk Menerapkan Style pada Shape
Private Sub ApplyStyle(ByRef shp As shape)
    shp.Fill.ApplyNoFill
    shp.Outline.width = mLineWidth / 10
    shp.Outline.Color.RGBAssign 0, 0, 0
End Sub

