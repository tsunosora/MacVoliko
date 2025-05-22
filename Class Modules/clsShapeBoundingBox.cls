VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShapeBoundingBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pUseOutline As Boolean

' -- Properti: untuk set atau get apakah ingin menyertakan outline
Public Property Let UseOutline(ByVal val As Boolean)
    pUseOutline = val
End Property

Public Property Get UseOutline() As Boolean
    UseOutline = pUseOutline
End Property

' -- Method utama: Buat bounding box per shape
Public Sub CreateBoundingBoxPerShape()
    Dim manager As New clsCorelSelectionManager
    Dim selRange As ShapeRange
    Dim shp As shape
    Dim outlineW As Double
    Dim leftX As Double, rightX As Double
    Dim topY As Double, bottomY As Double
    Dim rect As shape

    manager.UpdateSelection
    Set selRange = manager.SelectedShapes

    If selRange Is Nothing Or selRange.Count = 0 Then
        MsgBox "Tidak ada shape yang dipilih!", vbExclamation
        Exit Sub
    End If

    For Each shp In selRange
        ' Ambil nilai outline jika diminta
        If pUseOutline Then
            outlineW = shp.Outline.width
        Else
            outlineW = 0
        End If

        ' Hitung bounding box manual
        leftX = shp.leftX - (outlineW / 2)
        rightX = shp.rightX + (outlineW / 2)
        topY = shp.topY + (outlineW / 2)
        bottomY = shp.bottomY - (outlineW / 2)

        ' Buat rectangle sesuai bounding box
        Set rect = ActiveLayer.CreateRectangle(leftX, topY, rightX, bottomY)
        rect.Fill.ApplyNoFill
        rect.Outline.width = 0.3
        rect.Outline.Color.RGBAssign 255, 0, 0 ' Merah

    Next shp

End Sub


