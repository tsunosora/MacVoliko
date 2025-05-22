VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAutoDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === clsAutoDimension.cls ===
Option Explicit

Private targetShape As shape

' Set shape target dari luar
Public Sub SetShape(s As shape)
    Set targetShape = s
End Sub

' Fungsi umum untuk pengaturan teks
Private Sub SetTextProperties(dimShape As shape)
    With dimShape.Dimension.TextShape
        .text.Story.Font = "Arial"
        .text.Story.Size = 200
        .Fill.UniformColor.RGBAssign 41, 115, 175 ' Warna biru
    End With
End Sub

' Dimensi horizontal (lebar)
Public Sub DrawHorizontal()
    Dim dimShape As shape
    Dim snapPointLeft As SnapPoint
    Dim snapPointRight As SnapPoint
    Dim posX As Double, posY As Double
    Dim xS As Double ' Deklarasi xS sebagai variabel lokal

    ActiveDocument.Unit = cdrCentimeter

    ' Ambil snap points dari kiri ke kanan
    Set snapPointLeft = targetShape.SnapPoints.item(3) ' Snap kiri
    Set snapPointRight = targetShape.SnapPoints.item(4) ' Snap kanan

    xS = targetShape.SizeWidth / 10 ' Gunakan SizeWidth dari targetShape
    ' Tentukan posisi X dan Y untuk dimensi
    posX = targetShape.leftX + (targetShape.SizeWidth / 2)
    posY = targetShape.topY + xS ' Tempatkan teks sedikit di atas shape

    ' Buat garis dimensi horizontal
    Set dimShape = ActiveLayer.CreateLinearDimension( _
        cdrDimensionHorizontal, _
        snapPointLeft, _
        snapPointRight, _
        True, posX, posY, _
        Units:=cdrDimensionUnitCM)

    ' Modifikasi teks dimensi
    SetTextProperties dimShape
End Sub

' Dimensi vertikal (tinggi)
Public Sub DrawVertical()
    Dim dimShape As shape
    Dim snapPointTop As SnapPoint
    Dim snapPointBottom As SnapPoint
    Dim posX As Double, posY As Double
    Dim xS As Double ' Deklarasi xS sebagai variabel lokal

    ActiveDocument.Unit = cdrCentimeter

    xS = targetShape.SizeHeight / 10 ' Gunakan SizeHeight dari targetShape
    ' Ambil snap points dari atas ke bawah
    Set snapPointTop = targetShape.SnapPoints.BBox(cdrTopLeft) ' Snap atas
    Set snapPointBottom = targetShape.SnapPoints.BBox(cdrBottomLeft) ' Snap bawah

    ' Tentukan posisi X dan Y untuk dimensi
    posX = targetShape.leftX - xS ' Sedikit di kiri shape
    posY = targetShape.topY + (targetShape.SizeHeight / 2) ' Tengah vertikal shape

    ' Buat garis dimensi vertikal
    Set dimShape = ActiveLayer.CreateLinearDimension( _
        cdrDimensionVertical, _
        snapPointTop, _
        snapPointBottom, _
        True, posX, posY, _
        Units:=cdrDimensionUnitCM)

    ' Modifikasi teks dimensi
    SetTextProperties dimShape
End Sub


