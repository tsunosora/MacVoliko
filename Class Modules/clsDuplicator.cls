VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDuplicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: clsDuplicator
Public Rows As Integer
Public Cols As Integer
Public OffsetX As Double
Public OffsetY As Double
Public GridWidth As Double
Public GridHeight As Double
Public SpacingX As Double
Public SpacingY As Double
Private pUnit As cdrUnit

Public Property Let Unit(ByVal value As cdrUnit)
    pUnit = value
End Property

Public Property Get Unit() As cdrUnit
    Unit = pUnit
End Property

Public Sub Run()
    Dim doc As Document
    Dim SelectedShapes As ShapeRange
    Dim groupShape As shape
    Dim duplicateGroup As shape
    Dim startX As Double, startY As Double
    Dim i As Integer, j As Integer

    Set doc = ActiveDocument
    doc.Unit = Me.Unit

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Pilih satu atau lebih objek terlebih dahulu!", vbExclamation
        Exit Sub
    End If

    Set SelectedShapes = ActiveSelectionRange
    Set groupShape = SelectedShapes.Group

    startX = groupShape.leftX
    startY = groupShape.topY

    For i = 0 To Rows - 1
        For j = 0 To Cols - 1
            If Not (i = 0 And j = 0) Then
                Set duplicateGroup = groupShape.Duplicate
                duplicateGroup.SetPosition startX + (j * OffsetX), startY - (i * OffsetY)
            End If
        Next j
    Next i

    groupShape.UngroupAll

    MsgBox "Duplikasi selesai!", vbInformation
End Sub

' Mode otomatis: menghitung Rows & Cols berdasarkan ukuran objek + spacing
Public Sub RunAutomate()
    Dim doc As Document
    Dim selectedShape As shape
    Dim objWidth As Double, objHeight As Double
    Dim currentX As Double, currentY As Double
    Dim Rows As Integer, Cols As Integer
    Dim duplicateShape As shape
    Dim i As Integer, j As Integer

    Set doc = ActiveDocument
    doc.Unit = Me.Unit

    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Pilih satu objek terlebih dahulu!", vbExclamation
        Exit Sub
    End If

    Set selectedShape = ActiveSelectionRange(1)

    objWidth = selectedShape.SizeWidth
    objHeight = selectedShape.SizeHeight

    If objWidth > GridWidth Or objHeight > GridHeight Then
        MsgBox "Objek terlalu besar untuk ukuran grid yang diberikan.", vbExclamation
        Exit Sub
    End If

    If SpacingX < 0 Or SpacingY < 0 Then
        MsgBox "Jarak antar objek tidak boleh negatif!", vbExclamation
        Exit Sub
    End If

    ' Hitung jumlah baris dan kolom berdasarkan ukuran grid + spacing
    Rows = Int((GridHeight + SpacingY) / (objHeight + SpacingY))
    Cols = Int((GridWidth + SpacingX) / (objWidth + SpacingX))

    For i = 0 To Rows - 1
        For j = 0 To Cols - 1
            currentX = j * (objWidth + SpacingX)
            currentY = -i * (objHeight + SpacingY)

            If Not (i = 0 And j = 0) Then
                Set duplicateShape = selectedShape.Duplicate()
                duplicateShape.SetPosition currentX + selectedShape.leftX, currentY + selectedShape.topY
            End If
        Next j
    Next i
End Sub

Private Function GetUnitName(unitType As cdrUnit) As String
    Select Case unitType
        Case cdrMillimeter: GetUnitName = "mm"
        Case cdrCentimeter: GetUnitName = "cm"
        Case cdrInch: GetUnitName = "inch"
        Case cdrPixel: GetUnitName = "px"
        Case Else: GetUnitName = ""
    End Select
End Function

