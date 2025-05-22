VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuplicator 
   Caption         =   "UserForm1"
   ClientHeight    =   2670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "frmDuplicator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuplicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Nama : frmDuplicator
' \\\\\\\\\\\\\\\
' ===============(inialisasai userform)===============

Private Sub UserForm_Initialize()
    
    ' Set opsi default ke "Organize"
    optOrganize.value = True

    ' Set warna teks pada kontrol UserForm
    optOrganize.ForeColor = RGB(79, 77, 95)
    optAuto.ForeColor = RGB(79, 77, 95)
    cmdRun.ForeColor = RGB(79, 77, 95)
    cmdRunAuto.ForeColor = RGB(79, 95, 95)
    cmdExit.ForeColor = RGB(79, 77, 75)

    ' Set warna teks label
    Label1.ForeColor = RGB(79, 77, 95)
    Label2.ForeColor = RGB(79, 77, 95)
    Label3.ForeColor = RGB(79, 77, 95)
    Label4.ForeColor = RGB(79, 77, 95)
    Label5.ForeColor = RGB(79, 77, 95)
    Label6.ForeColor = RGB(79, 77, 95)
    Label7.ForeColor = RGB(79, 77, 95)
    Label8.ForeColor = RGB(79, 77, 95)
    Label9.ForeColor = RGB(79, 77, 95)

    ' Set ControlTipText
    cmbUnit.ControlTipText = "Document Unit..."
    txtRows.ControlTipText = "Baris..."
    txtCols.ControlTipText = "Kolom..."
    txtOffsetX.ControlTipText = "Jarak Horizontal dari posisi objek..."
    txtOffsetY.ControlTipText = "Jarak Vertical dari posisi objek..."
    txtGridWidth.ControlTipText = "Lebar Area duplikasi..."
    txtGridHeight.ControlTipText = "Tinggi Area duplikasi..."
    txtSpacingX.ControlTipText = "Jarak Horizontal antar objek..."
    txtSpacingY.ControlTipText = "Jarak Vertikal antar objek..."
    btnUndo.ControlTipText = "Undo / Reset hasil dari..."
    
    ' Inisialisasi nilai input awal
    txtRows.value = "0"
    txtCols.value = "0"
    txtSpacingX.value = "0"
    txtSpacingY.value = "0"

    ' Tambahkan pilihan satuan ke combobox
    With cmbUnit
        .AddItem "Millimeter"
        .AddItem "Centimeter"
        .AddItem "Inch"
        .AddItem "Pixel"
    End With

    ' Set default combobox berdasarkan unit dokumen saat ini
    Select Case ActiveDocument.Unit
        Case cdrMillimeter: cmbUnit.value = "Millimeter"
        Case cdrCentimeter: cmbUnit.value = "Centimeter"
        Case cdrInch: cmbUnit.value = "Inch"
        Case cdrPixel: cmbUnit.value = "Pixel"
        Case Else: cmbUnit.ListIndex = 0
    End Select

    ' (Duplikasi logika di atas, bisa disederhanakan, tapi sesuai permintaan tidak diubah)
    Select Case ActiveDocument.Unit
        Case cdrMillimeter: cmbUnit.value = "Millimeter"
        Case cdrCentimeter: cmbUnit.value = "Centimeter"
        Case cdrInch: cmbUnit.value = "Inch"
        Case cdrPixel: cmbUnit.value = "Pixel"
        Case Else: cmbUnit.ListIndex = 0
    End Select

    ' Set ulang unit dokumen saat form dibuka (men-trigger cmbUnit_Change)
    cmbUnit_Change

    ' Update ukuran grid berdasarkan unit
    UpdateGridSizeBasedOnUnit
End Sub

' \\\\\\\\\\\\\\\
' ===============(pilih satuan unit)===============
Private Sub cmbUnit_Change()
    ' Mengubah satuan dokumen sesuai pilihan user di combobox
    Select Case cmbUnit.value
        Case "Millimeter": ActiveDocument.Unit = cdrMillimeter
        Case "Centimeter": ActiveDocument.Unit = cdrCentimeter
        Case "Inch": ActiveDocument.Unit = cdrInch
        Case "Pixel": ActiveDocument.Unit = cdrPixel
    End Select

    ' Update ukuran grid setelah unit diubah
    UpdateGridSizeBasedOnUnit
End Sub

' \\\\\\\\\\\\\\\
' ===============(pilih mode #aktifasi tombol)===============
Private Sub optOrganize_Click()
    ' Aktifkan input untuk mode Organize
    txtRows.Enabled = True
    txtCols.Enabled = True
    txtOffsetX.Enabled = True
    txtOffsetY.Enabled = True
    cmdRun.Enabled = True
    Label2.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True

    ' Nonaktifkan input untuk mode Auto
    txtGridWidth.Enabled = False
    txtGridHeight.Enabled = False
    txtSpacingX.Enabled = False
    txtSpacingY.Enabled = False
    cmdRunAuto.Enabled = False
    Label6.Enabled = False
    Label7.Enabled = False
    Label8.Enabled = False
    Label9.Enabled = False
End Sub

' \\\\\\\\\\\\\\\
' ===============(pilih mode #aktifasi tombol)===============
Private Sub optAuto_Click()
    ' Aktifkan input untuk mode Auto
    txtGridWidth.Enabled = True
    txtGridHeight.Enabled = True
    txtSpacingX.Enabled = True
    txtSpacingY.Enabled = True
    cmdRunAuto.Enabled = True
    Label6.Enabled = True
    Label7.Enabled = True
    Label8.Enabled = True
    Label9.Enabled = True

    ' Nonaktifkan input untuk mode Organize
    txtRows.Enabled = False
    txtCols.Enabled = False
    txtOffsetX.Enabled = False
    txtOffsetY.Enabled = False
    cmdRun.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
End Sub

' \\\\\\\\\\\\\\\
' ===============(cmd perintah)===============
Private Sub cmdRun_Click()
    ' Inisialisasi objek duplicator manual (organize mode)
    Dim dup As New clsDuplicator
    
    ' Mulai grup command untuk menggabungkan semua aksi sebagai satu undo
    ActiveDocument.BeginCommandGroup "Buat Garis Otomatis"

    On Error GoTo errHandler
    
    ' Set unit yang dipilih user
    dup.Unit = GetSelectedUnit()

    ' Validasi dan ambil input jumlah baris, kolom, dan offset
    If IsNumeric(txtRows.value) And IsNumeric(txtCols.value) And _
       IsNumeric(txtOffsetX.value) And IsNumeric(txtOffsetY.value) Then
        dup.Rows = CInt(txtRows.value)
        dup.Cols = CInt(txtCols.value)
        dup.OffsetX = CDbl(txtOffsetX.value)
        dup.OffsetY = CDbl(txtOffsetY.value)
    Else
        MsgBox "Input grid tidak valid.", vbExclamation
        Exit Sub
    End If

    ' Jalankan proses duplikasi
    dup.Run
    
CleanUp:
    ' Akhiri grup command meskipun ada kesalahan input
    ActiveDocument.EndCommandGroup
    Exit Sub

errHandler:
End Sub

' \\\\\\\\\\\\\\\
' ===============(fungsi)===============
Private Function GetSelectedUnit() As cdrUnit
    ' Mengembalikan nilai unit sesuai pilihan user
    Select Case cmbUnit.value
        Case "Millimeter": GetSelectedUnit = cdrMillimeter
        Case "Centimeter": GetSelectedUnit = cdrCentimeter
        Case "Inch": GetSelectedUnit = cdrInch
        Case "Pixel": GetSelectedUnit = cdrPixel
        Case Else: GetSelectedUnit = cdrMillimeter ' default fallback
    End Select
End Function

' \\\\\\\\\\\\\\\
' ===============(cmd perintah)===============
Private Sub cmdRunAuto_Click()
    ' Inisialisasi objek duplicator otomatis (auto mode)
    Dim dup As New clsDuplicator

    ' Mulai grup command untuk menggabungkan semua aksi sebagai satu undo
    ActiveDocument.BeginCommandGroup "Buat Garis Otomatis"

    On Error GoTo errHandler
    
    ' Set unit yang dipilih
    dup.Unit = GetSelectedUnit()

    ' Validasi dan ambil input ukuran grid
    If IsNumeric(txtGridWidth.value) And IsNumeric(txtGridHeight.value) Then
        dup.GridWidth = CDbl(txtGridWidth.value)
        dup.GridHeight = CDbl(txtGridHeight.value)
    Else
        MsgBox "Input ukuran grid tidak valid.", vbExclamation
        Exit Sub
    End If

    ' Validasi dan ambil input jarak antar objek
    If IsNumeric(txtSpacingX.value) And IsNumeric(txtSpacingY.value) Then
        dup.SpacingX = CDbl(txtSpacingX.value)
        dup.SpacingY = CDbl(txtSpacingY.value)
    Else
        MsgBox "Jarak antar objek tidak valid.", vbExclamation
        Exit Sub
    End If

    ' Jalankan proses duplikasi otomatis
    dup.RunAutomate
    
CleanUp:
    ' Akhiri grup command meskipun ada kesalahan input
    ActiveDocument.EndCommandGroup
    Exit Sub

errHandler:
End Sub

' \\\\\\\\\\\\\\\
' ===============(Subrotin)===============
Private Sub UpdateGridSizeBasedOnUnit()
    ' Update nilai default grid berdasarkan satuan yang dipilih
    Select Case cmbUnit.value
        Case "Millimeter"
            txtGridWidth.value = 310
            txtGridHeight.value = 470
        Case "Centimeter"
            txtGridWidth.value = 31
            txtGridHeight.value = 47
        Case "Inch"
            txtGridWidth.value = 12.2
            txtGridHeight.value = 18.5
        Case "Pixel"
            txtGridWidth.value = 1172
            txtGridHeight.value = 1777
        Case Else
            txtGridWidth.value = 310
            txtGridHeight.value = 470
    End Select
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
