VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCorelSelectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: CorelSelectionManager
Private selShapes As ShapeRange

' Method untuk menyimpan shape yang terseleksi
Public Sub UpdateSelection()
    If Not ActiveDocument Is Nothing Then
        If ActiveDocument.Selection.Shapes.Count > 0 Then
            Set selShapes = ActiveDocument.Selection.Shapes.All
        Else
            Set selShapes = Nothing
        End If
    Else
        Set selShapes = Nothing
    End If
End Sub

' Property Get untuk mengambil ShapeRange
Public Property Get SelectedShapes() As ShapeRange
    Set SelectedShapes = selShapes
End Property

