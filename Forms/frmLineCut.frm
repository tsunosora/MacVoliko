VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLineCut 
   Caption         =   "UserForm1"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "frmLineCut.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLineCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' === frmWarnaPicker ===
Private WarnaPicker As clsWarnaPicker
Private ColorRGB As clsWarnaPicker


Private Sub UserForm_Initialize()
        
    btnCreate.ForeColor = RGB(79, 77, 95)
    cmdExit.ForeColor = RGB(79, 77, 95)
    Label9.ForeColor = RGB(79, 77, 95)
    Label10.ForeColor = RGB(79, 77, 95)
    
    Set WarnaPicker = New clsWarnaPicker
    Set ColorRGB = New clsWarnaPicker
    
    lblPreview.BackColor = RGB(0, 0, 0)
    
    ' Nilai default untuk panjang garis cut (dalam mm)
    txLineLength.value = "3"       ' Default 3 mm

    ' Nilai default untuk ketebalan garis (dalam mm)
    txtLineWidth.value = "0,2"     ' Default 0.2 mm (garis tipis)
    
End Sub

Private Sub btnCreate_Click()
    Dim cutter As New clsLineCutter

    ' Mulai grup command untuk menggabungkan semua aksi sebagai satu undo
    ActiveDocument.BeginCommandGroup "Buat Garis Otomatis"

    On Error GoTo errHandler

    ' Ambil warna dari preview
    cutter.lineColor = lblPreview.BackColor

    ' Validasi dan ambil nilai panjang garis
    If IsNumeric(txLineLength.value) Then
        cutter.lineLength = CDbl(txLineLength.value)
    Else
        MsgBox "Panjang garis tidak valid!", vbExclamation
        GoTo CleanUp
    End If

    ' Validasi dan ambil nilai ketebalan garis
    If IsNumeric(txtLineWidth.value) Then
        cutter.LineThickness = CDbl(txtLineWidth.value)
    Else
        MsgBox "Ketebalan garis tidak valid!", vbExclamation
        GoTo CleanUp
    End If

    ' Jalankan proses pemotongan garis
    cutter.Run

CleanUp:
    ' Akhiri grup command meskipun ada kesalahan input
    ActiveDocument.EndCommandGroup
    Exit Sub

errHandler:
    MsgBox "Terjadi kesalahan: " & Err.Description, vbCritical
    Resume CleanUp
End Sub


Private Sub cmdPilihWarna_Click()
    If WarnaPicker.PilihWarna Then
        ' Update tampilan preview warna
        lblPreview.BackColor = WarnaPicker.GetWarnaRGB
    Else
        MsgBox "Warna tidak dipilih.", vbInformation
    End If
End Sub

Private Sub btnUndo_Click()
    Call Undo
End Sub

Sub Undo()
    ActiveDocument.Undo 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

