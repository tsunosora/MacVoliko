VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLabelsBanner 
   Caption         =   "UserForm2"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6495
   OleObjectBlob   =   "frmLabelsBanner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLabelsBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' \\\\\\\\\\\\\\\
' ===============(Inialisai userform)===============
Private Sub UserForm_Initialize()
    Dim i As Integer
    
    ' Load dulu form sumber, kalau belum dimuat
    Load frmDefaultSetting
    
    For i = 0 To frmDefaultSetting.cmbDefaultBahan.ListCount - 1
        cmbBahan.AddItem frmDefaultSetting.cmbDefaultBahan.List(i)
    Next i
    
    ' Tambahkan item ke cmbAlamat
    With cmbAlamat
        .AddItem "Sew"
        .AddItem "Imo"
    End With
        
    ' Set warna teks pada kontrol UserForm
    LabelNm.ForeColor = RGB(79, 77, 95)
    LabelQty.ForeColor = RGB(79, 77, 95)
    LabelFYI.ForeColor = RGB(79, 77, 95)
    LabelMat.ForeColor = RGB(79, 77, 95)
    chkDate.ForeColor = RGB(79, 77, 95)
    chkTempat.ForeColor = RGB(79, 77, 95)
    chkExpress.ForeColor = RGB(79, 77, 95)
    cmdCreateLabels.ForeColor = RGB(79, 77, 95)
    cmdGetFinishing.ForeColor = RGB(79, 77, 95)
    cmdExit.ForeColor = RGB(79, 77, 95)
    
End Sub

' \\\\\\\\\\\\\\\
' ===============(BERALIH KE MODE SETTING DEFAULT)===============
Private Sub CommandButton1_Click()
    frmDefaultSetting.Show VBMODELES
End Sub

' \\\\\\\\\\\\\\\
' ===============(inialisasi userform)===============
Private Sub cmdCreateLabels_Click()
    Dim lbl As New clsBannerLabel
    Dim originalSelection As ShapeRange
    Dim SR As ShapeRange, s As shape, s1 As shape
    Dim xPos As Double, yPos As Double

    ' Mulai grup command untuk menggabungkan semua aksi sebagai satu undo
    ActiveDocument.BeginCommandGroup "Buat Garis Otomatis"

    On Error GoTo errHandler
    
    ActiveDocument.Unit = cdrMeter
    
    ' Ambil data dari form
    lbl.CustomerName = txtNama.value
    lbl.Material = cmbBahan.value
    lbl.Finishing = txtKeteranagan.value
    lbl.Quantity = txtJumlah.value
    lbl.Address = cmbAlamat.value
    lbl.ShowDate = chkDate.value
    lbl.ShowAddress = chkTempat.value
    lbl.IsExpress = chkExpress.value

    Set SR = ActiveSelectionRange
    If SR.Count = 0 Then
        MsgBox "No shapes selected!", vbExclamation
        Exit Sub
    End If
    
    ' Simpan seleksi awal
    Set originalSelection = ActiveSelectionRange
    
    For Each s In SR
        ' Hitung posisi label: 3 cm dari kiri dan 1.5 cm di atas
        xPos = s.leftX + 0.03 ' 3 cm = 0.03 meter
        yPos = s.topY + 0.011 ' 1.11 cm = 0.011 meter

        ' Buat teks di posisi tersebut
        Set s1 = ActiveLayer.CreateArtisticText(xPos, yPos, lbl.GetLabelText(s.SizeWidth, s.SizeHeight))

        ' Ubah warna teks jadi merah
        s1.Fill.UniformColor.RGBAssign 255, 0, 0
    Next s
         ' Seleksi ulang shape yang dipilih di awal
        ActiveDocument.ClearSelection
        originalSelection.CreateSelection
CleanUp:
    ' Akhiri grup command meskipun ada kesalahan input
    ActiveDocument.EndCommandGroup
    Exit Sub

errHandler:
    MsgBox "Terjadi kesalahan: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' \\\\\\\\\\\\\\\
' ===============(Undo)===============
Private Sub btnUndo_Click()
    Call Undo
End Sub

Sub Undo()
    ActiveDocument.Undo 1
End Sub

' \\\\\\\\\\\\\\\
' ===============(Close)===============
Private Sub cmdExit_Click()
    Unload Me
End Sub

' \\\\\\\\\\\\\\\
' ===============(beralih mode)===============
Private Sub cmdGetFinishing_Click()
    frmFinishing.Show VBMODELES
    Unload Me
End Sub
