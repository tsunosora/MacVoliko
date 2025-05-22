VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bmpManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: bmpManager

' ============ Set Class ============
Private bmpClassic As shape
Private bmpModern As shape
Private bmpSilver As shape
Private bmpGold As shape

' ============ Properti Public ============
Public Property Get rivetClassic() As shape
    Set rivetClassic = bmpClassic
End Property

Public Property Get rivetModern() As shape
    Set rivetModern = bmpModern
End Property

Public Property Get rivetSilver() As shape
    Set rivetSilver = bmpSilver
End Property

Public Property Get rivetGold() As shape
    Set rivetGold = bmpGold
End Property

' ============ Subrotin ============
Sub GenerateRivetClassic()
    ' === Deklarasi ===
    Dim rec1 As shape ' Persegi
    Dim elps1 As shape ' Elips
    Dim sGroup As shape
    Dim rectSize As Double: rectSize = 2.5      ' Membuat nilai
    Dim ellipseSize As Double: ellipseSize = 1  ' Membuat nilai

    ' === Menentukan unit perintah ===
    ActiveDocument.Unit = cdrCentimeter

    ' === Buat persegi ===
    Set rec1 = ActiveLayer.CreateRectangle2(0, 0, rectSize, rectSize)
    rec1.rectangle.CornerType = cdrCornerTypeRound
    rec1.rectangle.RelativeCornerScaling = True
    rec1.Fill.ApplyNoFill
    rec1.Outline.SetPropertiesEx 0.0002, OutlineStyles(0), CreateRGBColor(79, 77, 95), _
        ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
        cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, _
        Justification:=cdrOutlineJustificationMiddle
    ' Menghilangkan outline persegi
    rec1.Outline.SetNoOutline

    ' === Buat ellipse (tengah persegi) ===
    Set elps1 = ActiveLayer.CreateEllipse2(rectSize / 2, rectSize / 2, ellipseSize / 2, -ellipseSize / 2)
    elps1.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    elps1.Outline.SetPropertiesEx 0.075, OutlineStyles(0), CreateRGBColor(79, 77, 95), _
        ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
        cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, _
        Justification:=cdrOutlineJustificationMiddle

    ' === Group semua shape ===
    Set sGroup = ActiveDocument.CreateShapeRangeFromArray(rec1, elps1).Group
    ' === Seleksi shape ketika perintah selesai ===
    sGroup.Selected = True
    
    ' === Seleksi aktif shape ===
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    ' === Convert shape ke BMP ===
    Dim classicBit As shape
    Set classicBit = OrigSelection.ConvertToBitmapEx(5, False, True, 300, 1, True, False, 95)
    Set bmpClassic = classicBit
End Sub

' ============ Subrotin ============
Sub GenerateRivetModern()
    ' === Deklarasi ===
    Dim recB As shape          ' Persegi urtama
    Dim rectSize As Double: rectSize = 2.5
    Dim elps1 As shape         ' Elips utama
    Dim rec1 As shape          ' Persegi kiri
    Dim rec2 As shape          ' Persegi kanan (mirror)
    Dim groupedRects As shape  ' Grup dua persegi
    Dim rectWidth As Double: rectWidth = 0.09           ' Membuat nilai
    Dim rectHeight As Double: rectHeight = 0.8          ' Membuat nilai
    Dim ellipseDiameter As Double: ellipseDiameter = 1# ' Membuat nilai
    Dim sGroup As shape
    
    ' === Menentukan unit perintah ===
    ActiveDocument.Unit = cdrCentimeter
    
    ' === Buat persegi ===
    Set recB = ActiveLayer.CreateRectangle2(0, 0, rectSize, rectSize)
    recB.rectangle.CornerType = cdrCornerTypeRound
    recB.rectangle.RelativeCornerScaling = True
    recB.Fill.ApplyNoFill
    recB.Outline.SetPropertiesEx 0.0002, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), _
        ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
        cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, _
        Justification:=cdrOutlineJustificationMiddle
    ' Menghilangkan outline persegi
    recB.Outline.SetNoOutline
    
    ' === Buat Elips 1cm diameter (frame luar) ===
    Set elps1 = ActiveLayer.CreateEllipse2(0, 0, ellipseDiameter / 2, -ellipseDiameter / 2)
    With elps1
        .Fill.UniformColor.CMYKAssign 0, 0, 0, 0
        .Outline.SetPropertiesEx 0.075, OutlineStyles(0), CreateRGBColor(79, 77, 95), _
            ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
            cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, _
            MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    End With

    ' === Buat Persegi Merah Pertama ===
    Set rec1 = ActiveLayer.CreateRectangle2(0, 0, rectWidth, rectHeight)
    With rec1
        .rectangle.CornerType = cdrCornerTypeRound
        .rectangle.RelativeCornerScaling = True
        .Fill.UniformColor.CMYKAssign 0, 100, 100, 0 ' Merah
        .Outline.SetNoOutline
        .Rotate 45
    End With

    ' === Duplikat dan Mirror ===
    rec1.Copy
    ActiveLayer.Paste
    Set rec2 = ActiveSelection
    rec2.Flip cdrFlipHorizontal
    'rec2.Rotate 300 ' Untuk tetap menyilang

    ' === Distribute ke Tengah Horizontal ===
    rec1.AddToSelection
    ActiveSelectionRange.Distribute cdrDistributeCenter, cdrHorizontal

    ' === Group dan Masukkan ke Elips ===
    Set groupedRects = ActiveSelection.Group
    groupedRects.AddToPowerClip elps1, -1
    
    ' === Distribute shape array ===
    ActiveDocument.CreateShapeRangeFromArray(recB, elps1).AlignAndDistribute 3, 3, 0, 0, False, 2
    
    ' === Group semua shape ===
    Set sGroup = ActiveDocument.CreateShapeRangeFromArray(recB, elps1).Group
    ' === Seleksi shape ketika perintah selesai ===
    sGroup.Selected = True

    ' === Seleksi aktif shape ===
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    ' === Convert shape ke BMP ===
    Dim modernBit As shape
    Set modernBit = OrigSelection.ConvertToBitmapEx(5, False, True, 300, 1, True, False, 95)
    Set bmpModern = modernBit
End Sub

' ============ Subrotin ============
Sub GenerateRivetSilver()
    ' === Deklarasi ===
    Dim rec1 As shape ' Persegi
    Dim elps1 As shape         ' Elips utama
    Dim styleString As String
    Dim elps2 As shape         ' Elips primer
    Dim sGroup As shape
    Dim rectSize As Double: rectSize = 2.5      ' Membuat nilai
    Dim ellipseSize As Double: ellipseSize = 1  ' Membuat nilai

    ' === Menentukan unit perintah ===
    ActiveDocument.Unit = cdrCentimeter
    
    ' === Buat persegi ===
    Set rec1 = ActiveLayer.CreateRectangle2(0, 0, rectSize, rectSize)
    rec1.rectangle.CornerType = cdrCornerTypeRound
    rec1.rectangle.RelativeCornerScaling = True
    rec1.Fill.ApplyNoFill
    rec1.Outline.SetPropertiesEx 0.0002, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), _
        ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
        cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, _
        Justification:=cdrOutlineJustificationMiddle
    ' Menghilangkan outline persegi
    rec1.Outline.SetNoOutline
    
    ' === Membuat elips utama ===
    Set elps1 = ActiveLayer.CreateEllipse2(0, 0, ellipseSize / 2, ellipseSize / 2)
    
    ' Menghapus fill dan mengatur outline
    elps1.Fill.ApplyNoFill
    elps1.Outline.SetPropertiesEx 0.02, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 0), _
                              ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, _
                              cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    ' Menyusun string gaya yang digunakan berulang
    styleString = "{""StackedBitmapEffects"":{},""fill"":{""fillName"":""2a03ecd2-88ec-43c5-a592-bb92c49275c9"",""intermediateColors"":[""49:RGB255,USER,209,209,209,100,00000000-0000-0000-0000-000000000000:255:0:50""],""mode"":""0"",""primaryColor"":""RGB255,USER,161,161,161,100,00000000-0000-0000-0000-000000000000"",""rateValue"":""50"",""secondaryColor"":""RGB255,USER,161,161,161,100,00000000-0000-0000-0000-000000000000"",""type"":""2""},""outline"":{""color"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000"",""width"":""2000""},""transparency"":{}}"
    
    ' Menerapkan gaya ke shape s1
    Dim i As Integer
    For i = 1 To 10
        elps1.Style.StringAssign styleString
    Next i
    
    ' Memutar shape
    elps1.Rotate 43
    ' Menghapus outline
    elps1.Outline.SetNoOutline
    
    ' === Membuat elips sekunder ===
    Set elps2 = ActiveLayer.CreateEllipse2(0, 0, ellipseSize / 5, ellipseSize / 5)
    
    ' Mengatur fill dan menghapus outline
    elps2.Fill.ApplyNoFill
    elps2.Outline.SetPropertiesEx 0.02, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), _
                              ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, _
                              cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    elps2.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    ' Menghapus outline
    elps2.Outline.SetNoOutline
                              
    ' === Distribute shape array ===
    ActiveDocument.CreateShapeRangeFromArray(rec1, elps1, elps2).AlignAndDistribute 3, 3, 0, 0, False, 2
    
    ' === Group semua shape ===
    Set sGroup = ActiveDocument.CreateShapeRangeFromArray(rec1, elps1, elps2).Group
    
    ' === Seleksi shape ketika perintah selesai ===
    sGroup.Selected = True

    ' === Seleksi aktif shape ===
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    ' === Convert shape ke BMP ===
    Dim silverBit As shape
    Set silverBit = OrigSelection.ConvertToBitmapEx(5, False, True, 300, 1, True, False, 95)
    Set bmpSilver = silverBit
End Sub

' ============ Subrotin ============
Sub GenerateRivetGold()
    ' === Deklarasi ===
    Dim rec1 As shape ' Persegi
    Dim elps1 As shape         ' Elips utama
    Dim styleString As String
    Dim elps2 As shape         ' Elips primer
    Dim sGroup As shape
    Dim rectSize As Double: rectSize = 2.5      ' Membuat nilai
    Dim ellipseSize As Double: ellipseSize = 1  ' Membuat nilai

    ' === Menentukan unit perintah ===
    ActiveDocument.Unit = cdrCentimeter
    
    ' === Buat persegi ===
    Set rec1 = ActiveLayer.CreateRectangle2(0, 0, rectSize, rectSize)
    rec1.rectangle.CornerType = cdrCornerTypeRound
    rec1.rectangle.RelativeCornerScaling = True
    rec1.Fill.ApplyNoFill
    rec1.Outline.SetPropertiesEx 0.0002, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), _
        ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, _
        cdrOutlineButtLineCaps, cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, _
        Justification:=cdrOutlineJustificationMiddle
    ' Menghilangkan outline persegi
    rec1.Outline.SetNoOutline
    
    ' === Membuat elips utama ===
    Set elps1 = ActiveLayer.CreateEllipse2(0, 0, ellipseSize / 2, ellipseSize / 2)
    
    ' Menghapus fill dan mengatur outline
    elps1.Fill.ApplyNoFill
    elps1.Outline.SetPropertiesEx 0.02, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 0), _
                              ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, _
                              cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    ' Menyusun string gaya yang digunakan berulang
    styleString = "{""StackedBitmapEffects"":{},""fill"":{""fillName"":""2a03ecd2-88ec-43c5-a592-bb92c49275c9"",""intermediateColors"":[""49:RGB255,USER,255,233,87,100,00000000-0000-0000-0000-000000000000:255:0:50""],""mode"":""0"",""primaryColor"":""RGB255,USER,255,174,0,100,00000000-0000-0000-0000-000000000000"",""rateValue"":""50"",""secondaryColor"":""RGB255,USER,255,174,0,100,00000000-0000-0000-0000-000000000000"",""type"":""2""},""outline"":{""color"":""CMYK,USER,0,0,0,100,100,00000000-0000-0000-0000-000000000000"",""width"":""2000""},""transparency"":{}}"
    
    ' Menerapkan gaya ke shape s1
    Dim i As Integer
    For i = 1 To 10
        elps1.Style.StringAssign styleString
    Next i
    
    ' Memutar shape
    elps1.Rotate 43
    ' Menghapus outline
    elps1.Outline.SetNoOutline
    
    ' === Membuat elips sekunder ===
    Set elps2 = ActiveLayer.CreateEllipse2(0, 0, ellipseSize / 5, ellipseSize / 5)
    
    ' Mengatur fill dan menghapus outline
    elps2.Fill.ApplyNoFill
    elps2.Outline.SetPropertiesEx 0.02, OutlineStyles(0), CreateCMYKColor(0, 0, 0, 100), _
                              ArrowHeads(0), ArrowHeads(0), cdrFalse, cdrFalse, cdrOutlineButtLineCaps, _
                              cdrOutlineMiterLineJoin, 0#, 100, MiterLimit:=5#, Justification:=cdrOutlineJustificationMiddle
    elps2.Fill.UniformColor.CMYKAssign 0, 0, 0, 0
    ' Menghapus outline
    elps2.Outline.SetNoOutline
                              
    ' === Distribute shape array ===
    ActiveDocument.CreateShapeRangeFromArray(rec1, elps1, elps2).AlignAndDistribute 3, 3, 0, 0, False, 2
    
    ' === Group semua shape ===
    Set sGroup = ActiveDocument.CreateShapeRangeFromArray(rec1, elps1, elps2).Group
    
    ' === Seleksi shape ketika perintah selesai ===
    sGroup.Selected = True

    ' === Seleksi aktif shape ===
    Dim OrigSelection As ShapeRange
    Set OrigSelection = ActiveSelectionRange
    
    ' === Convert shape ke BMP ===
    Dim goldBit As shape
    Set goldBit = OrigSelection.ConvertToBitmapEx(5, False, True, 300, 1, True, False, 95)
    Set bmpGold = goldBit
End Sub





