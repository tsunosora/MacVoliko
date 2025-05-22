Attribute VB_Name = "ringOne_String"
Sub ringOne(jumlahAtas As Integer, jumlahBawah As Integer, jumlahKiri As Integer, jumlahKanan As Integer, rivetType As String)
    Dim originalSelection As ShapeRange
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Silakan pilih satu atau beberapa shape terlebih dahulu!"
        Exit Sub
    End If

    ' Simpan seleksi awal
    Set originalSelection = ActiveSelectionRange

    Dim s As shape
    For Each s In originalSelection
        Dim rivetGen As New RivetGenerator
        rivetGen.SetInputValues jumlahAtas, jumlahBawah, jumlahKiri, jumlahKanan, s, 2.5
        rivetGen.SetRivetType rivetType
        rivetGen.GenerateRivet
    Next s

    ' Seleksi ulang shape yang dipilih di awal
    ActiveDocument.ClearSelection
    originalSelection.CreateSelection
End Sub

