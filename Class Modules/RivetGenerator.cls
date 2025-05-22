VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RivetGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === Class Module: RivetGenerator ===
Private leftX As Double, rightX As Double, topY As Double, bottomY As Double
Private selectedShapeWidth As Double, selectedShapeHeight As Double
Private jumlahAtas As Integer, jumlahBawah As Integer
Private jumlahKiri As Integer, jumlahKanan As Integer
Private rivetWidth As Double
Private rivetGroup As New ShapeRange
Private rivetType As String
Private originalRivet As shape

' === Set posisi & ukuran shape utama ===
Public Sub SetInputValues(ByVal atas As Integer, ByVal bawah As Integer, ByVal kiri As Integer, ByVal kanan As Integer, _
                          ByVal shape As shape, Optional ByVal width As Double = 2.5)
    jumlahAtas = atas
    jumlahBawah = bawah
    jumlahKiri = kiri
    jumlahKanan = kanan
    selectedShapeWidth = shape.SizeWidth
    selectedShapeHeight = shape.SizeHeight
    leftX = shape.leftX
    rightX = shape.rightX
    topY = shape.topY
    bottomY = shape.bottomY
    rivetWidth = width
End Sub

Public Sub SetRivetType(ByVal tipe As String)
    rivetType = tipe
End Sub

Private Sub GenerateBaseRivet()
    Dim rivet As New bmpManager
    Select Case rivetType
        Case "Modern": rivet.GenerateRivetModern: Set originalRivet = rivet.rivetModern
        Case "Silver": rivet.GenerateRivetSilver: Set originalRivet = rivet.rivetSilver
        Case "Gold": rivet.GenerateRivetGold: Set originalRivet = rivet.rivetGold
        Case Else: rivet.GenerateRivetClassic: Set originalRivet = rivet.rivetClassic
    End Select

    originalRivet.SizeWidth = rivetWidth
    originalRivet.SizeHeight = rivetWidth
End Sub

Public Sub GenerateRivet()
    Dim i As Integer
    Dim rClassic As shape
    Dim posX As Double, posY As Double
    Dim totalSpacing As Double, spacing As Double
    Dim tempRange As ShapeRange
    Dim s As shape

    ' === Buat satu base rivet ===
    GenerateBaseRivet

    ' === ATAS ===
    If jumlahAtas > 0 Then
        If jumlahAtas = 1 Then
            posX = leftX + selectedShapeWidth / 2
            Set rClassic = originalRivet.Duplicate
            With rClassic
                .centerX = posX: .topY = topY
            End With
            rivetGroup.Add rClassic
        Else
            totalSpacing = selectedShapeWidth - (jumlahAtas * rivetWidth)
            If totalSpacing < 0 Then MsgBox "Rivet atas terlalu banyak!": Exit Sub
            spacing = totalSpacing / (jumlahAtas - 1)
            For i = 0 To jumlahAtas - 1
                posX = leftX + (i * (rivetWidth + spacing)) + rivetWidth / 2
                Set rClassic = originalRivet.Duplicate
                With rClassic
                    .centerX = posX: .topY = topY
                End With
                rivetGroup.Add rClassic
            Next i
        End If
    End If

    ' === BAWAH ===
    If jumlahBawah > 0 Then
        If jumlahBawah = 1 Then
            posX = leftX + selectedShapeWidth / 2
            Set rClassic = originalRivet.Duplicate
            With rClassic
                .centerX = posX: .bottomY = bottomY
            End With
            rivetGroup.Add rClassic
        Else
            totalSpacing = selectedShapeWidth - (jumlahBawah * rivetWidth)
            If totalSpacing < 0 Then MsgBox "Rivet bawah terlalu banyak!": Exit Sub
            spacing = totalSpacing / (jumlahBawah - 1)
            For i = 0 To jumlahBawah - 1
                posX = leftX + (i * (rivetWidth + spacing)) + rivetWidth / 2
                Set rClassic = originalRivet.Duplicate
                With rClassic
                    .centerX = posX: .bottomY = bottomY
                End With
                rivetGroup.Add rClassic
            Next i
        End If
    End If

    ' === KIRI ===
    If jumlahKiri > 0 Then
        jumlahKiri = jumlahKiri + 2
        totalSpacing = selectedShapeHeight - (jumlahKiri * rivetWidth)
        If totalSpacing < 0 Then MsgBox "Rivet kiri terlalu banyak!": Exit Sub
        spacing = totalSpacing / (jumlahKiri - 1)
        Set tempRange = New ShapeRange
        For i = 0 To jumlahKiri - 1
            posY = bottomY + (i * (rivetWidth + spacing)) + rivetWidth / 2
            Set rClassic = originalRivet.Duplicate
            With rClassic
                .CenterY = posY: .leftX = leftX
            End With
            tempRange.Add rClassic
        Next i
        tempRange(1).Delete
        tempRange(tempRange.Count).Delete
        For i = 1 To tempRange.Count
            On Error Resume Next
            Set s = tempRange(i)
            If Not s Is Nothing Then rivetGroup.Add s
            On Error GoTo 0
        Next i
    End If

    ' === KANAN ===
    If jumlahKanan > 0 Then
        jumlahKanan = jumlahKanan + 2
        totalSpacing = selectedShapeHeight - (jumlahKanan * rivetWidth)
        If totalSpacing < 0 Then MsgBox "Rivet kanan terlalu banyak!": Exit Sub
        spacing = totalSpacing / (jumlahKanan - 1)
        Set tempRange = New ShapeRange
        For i = 0 To jumlahKanan - 1
            posY = bottomY + (i * (rivetWidth + spacing)) + rivetWidth / 2
            Set rClassic = originalRivet.Duplicate
            With rClassic
                .CenterY = posY: .rightX = rightX
            End With
            tempRange.Add rClassic
        Next i
        tempRange(1).Delete
        tempRange(tempRange.Count).Delete
        For i = 1 To tempRange.Count
            On Error Resume Next
            Set s = tempRange(i)
            If Not s Is Nothing Then rivetGroup.Add s
            On Error GoTo 0
        Next i
    End If

    ' === Hapus Base Rivet ===
    If Not originalRivet Is Nothing Then
        originalRivet.Delete
    End If

    ' === GROUP ===
    If rivetGroup.Count > 0 Then
        rivetGroup.Group
    Else
        MsgBox "Tidak ada rivet yang dibuat."
    End If
End Sub


