VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBahanManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: clsBahanManager
Private bahanList As Collection
Private Const REG_KEY As String = "LabelsApp"
Private Const REG_SECTION As String = "UserData"
Private Const REG_VALUE As String = "DaftarBahan"

' Default list jika registry kosong
Private Function DefaultList() As Collection
    Dim col As New Collection
    col.Add "F280"
    col.Add "F340"
    col.Add "F380"
    col.Add "F440"
    col.Add "BACKLITE"
    col.Add "VINYL Indoor"
    Set DefaultList = col
End Function

Private Sub Class_Initialize()
    Set bahanList = New Collection
    LoadBahan
End Sub

Private Sub LoadBahan()
    On Error Resume Next
    Dim str As String: str = GetSetting(REG_KEY, REG_SECTION, REG_VALUE, "")
    
    If str = "" Then
        Set bahanList = DefaultList()
    Else
        Dim arr() As String: arr = Split(str, "|")
        Dim item As Variant
        For Each item In arr
            If Trim(item) <> "" Then
                bahanList.Add Trim(item)
            End If
        Next
    End If
End Sub

Public Function AmbilSemuaBahan() As Collection
    Set AmbilSemuaBahan = bahanList
End Function

Public Sub TambahBahan(bahan As String)
    Dim item As Variant
    For Each item In bahanList
        If StrComp(item, bahan, vbTextCompare) = 0 Then Exit Sub
    Next
    bahanList.Add Trim(bahan)
    SimpanKeRegistry
End Sub

Public Sub HapusBahan(bahan As String)
    Dim i As Integer
    For i = 1 To bahanList.Count
        If StrComp(bahanList(i), bahan, vbTextCompare) = 0 Then
            bahanList.Remove i
            Exit For
        End If
    Next
    SimpanKeRegistry
End Sub

Private Sub SimpanKeRegistry()
    Dim s As String, item As Variant
    For Each item In bahanList
        s = s & item & "|"
    Next
    SaveSetting REG_KEY, REG_SECTION, REG_VALUE, s
End Sub

