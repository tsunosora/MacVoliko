VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFinishing 
   Caption         =   "UserForm1"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8175
   OleObjectBlob   =   "frmFinishing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFinishing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' \\\\\\\\\\\\\\\
' ===============(inialisasi userform)===============
Private Sub UserForm_Initialize()
    ' Inisialisasi value yang di perlukan
    txtAtas.value = 2
    txtBawah.value = 2
    txtKiri.value = 1
    txtKanan.value = 1
    optClassic.value = True
    
    ' Sembunyikan
    HideAllButtons
    HideAllValue
    imgClassic.Visible = True
    
    ' Set warna teks
    optTunnelAB.ForeColor = RGB(79, 77, 95)
    optTunnelRL.ForeColor = RGB(79, 77, 95)
    optOver.ForeColor = RGB(79, 77, 95)
    optFitobject.ForeColor = RGB(79, 77, 95)
    optRivet.ForeColor = RGB(79, 77, 95)
    optClassic.ForeColor = RGB(79, 77, 95)
    optModern.ForeColor = RGB(79, 77, 95)
    optSilver.ForeColor = RGB(79, 77, 95)
    optGold.ForeColor = RGB(79, 77, 95)
    Labelline.ForeColor = RGB(79, 77, 95)
    optBounding.ForeColor = RGB(79, 77, 95)
    cmdExit.ForeColor = RGB(79, 77, 95)
    cmdGetLabels.ForeColor = RGB(79, 77, 95)
    cmdMakeFinishing.ForeColor = RGB(79, 77, 95)
    cmdMeasure.ForeColor = RGB(79, 77, 95)
    
    ' Set ControlTipText teks
    optTunnelAB.ControlTipText = "Kolong atas - bawah.."
    optTunnelRL.ControlTipText = "Kolong kanan - kiri.."
    optOver.ControlTipText = "Sisa bahan.."
    optFitobject.ControlTipText = "sisa bahan kecil, pendukung Rivet.."
    optRivet.ControlTipText = "Rivet / Keling.."
    txtLine.ControlTipText = "Ketebalan line.."
    optBounding.ControlTipText = "Menyertakan garis tepi.."
    cmdMeasure.ControlTipText = "Ukur object.."
    cmdMakeFinishing.ControlTipText = "Buat Finishing.."
    cmdGetLabels.ControlTipText = "Beralih ke mode Labels.."
    cmdExit.ControlTipText = "Keluar.."
    btnUndo.ControlTipText = "Undo.."
    imgClassic.ControlTipText = "Jenis rivet yang di pilih.."
    imgModern.ControlTipText = "Jenis rivet yang di pilih.."
    imgSilver.ControlTipText = "Jenis rivet yang di pilih.."
    imgGold.ControlTipText = "Jenis rivet yang di pilih.."
    
End Sub

' \\\\\\\\\\\\\\\
' ===============(mengatur pengaktifan value)===============
Private Sub optFitobject_Click()
    HideAllValue
    txtFit.Enabled = True
    optClassic.Enabled = False
    optModern.Enabled = False
    optSilver.Enabled = False
    optGold.Enabled = False
End Sub

Private Sub optOver_Click()
    HideAllValue
    txtOver.Enabled = True
    optClassic.Enabled = False
    optModern.Enabled = False
    optSilver.Enabled = False
    optGold.Enabled = False
End Sub

Private Sub optRivet_Click()
    HideAllValue
    txtAtas.Enabled = True
    txtBawah.Enabled = True
    txtKanan.Enabled = True
    txtKiri.Enabled = True
    optClassic.Enabled = True
    optModern.Enabled = True
    optSilver.Enabled = True
    optGold.Enabled = True
End Sub

Private Sub optTunnelAB_Click()
    HideAllValue
    txtTunnel.Enabled = True
    optClassic.Enabled = False
    optModern.Enabled = False
    optSilver.Enabled = False
    optGold.Enabled = False
End Sub

Private Sub optTunnelRL_Click()
    HideAllValue
    txtTunnel.Visible = True
    optClassic.Enabled = False
    optModern.Enabled = False
    optSilver.Enabled = False
    optGold.Enabled = False
End Sub

Private Sub optClassic_Click()
    HideAllButtons
    imgClassic.Visible = True
End Sub

Private Sub optModern_Click()
    HideAllButtons
    imgModern.Visible = True
End Sub

Private Sub optSilver_Click()
    HideAllButtons
    imgSilver.Visible = True
End Sub

Private Sub optGold_Click()
    HideAllButtons
    imgGold.Visible = True
End Sub
' hide semua value
Private Sub HideAllValue()
    txtTunnel.Enabled = False
    txtOver.Enabled = False
    txtFit.Enabled = False
    txtAtas.Enabled = False
    txtBawah.Enabled = False
    txtKanan.Enabled = False
    txtKiri.Enabled = False
    
    optClassic.Enabled = False
    optModern.Enabled = False
    optSilver.Enabled = False
    optGold.Enabled = False
End Sub
' set hide gambar
Private Sub HideAllButtons()
    imgClassic.Visible = True
    imgModern.Visible = False
    imgSilver.Visible = False
    imgGold.Visible = False
End Sub

' \\\\\\\\\\\\\\\
' ===============(Jalankan perintah)===============
Private Sub cmdMakeFinishing_Click()
    ' Deklarasi objek logic di awal
    Dim logic As New clsBannerFinishingLogic
    Dim boxer As New clsShapeBoundingBox
    
    ' Mulai grup command untuk menggabungkan semua aksi sebagai satu undo
    ActiveDocument.BeginCommandGroup "Buat Garis Otomatis"

    On Error GoTo errHandler
    
    ' Kondisi untuk optRivet
    If optRivet.value Then
        Dim jumlahAtas As Integer, jumlahBawah As Integer
        Dim jumlahKiri As Integer, jumlahKanan As Integer
        Dim rivetType As String

        ' Mengambil nilai dari textboxes
        jumlahAtas = val(txtAtas.value)
        jumlahBawah = val(txtBawah.value)
        jumlahKiri = val(txtKiri.value)
        jumlahKanan = val(txtKanan.value)

        ' Menentukan jenis rivet yang dipilih
        If optModern.value Then
            rivetType = "Modern"
        ElseIf optSilver.value Then
            rivetType = "Silver"
        ElseIf optGold.value Then
            rivetType = "Gold"
        Else
            rivetType = "Classic"
        End If

        ' Memanggil fungsi untuk rivet
        Call ringOne(jumlahAtas, jumlahBawah, jumlahKiri, jumlahKanan, rivetType)

    ' Kondisi untuk optOver / optFitobject
    ElseIf optOver.value Or optFitobject.value Then
        ' Mengisi properti dari class logic
        logic.FitSize = val(Me.txtFit.text)
        logic.OverSize = val(Me.txtOver.text)
        logic.UseOver = Me.optOver.value
        logic.LineWidth = val(Me.txtLine.text)

        ' Memanggil metode ApplySpacing dari class logic
        logic.ApplySpacing

    ' Kondisi untuk optTunnelAB / optTunnelRL
    ElseIf optTunnelAB.value Or optTunnelRL.value Then
        ' Mengisi properti dari class logic untuk Tunnel
        logic.TunnelSize = val(Me.txtTunnel.text)
        logic.LineWidth = val(Me.txtLine.text)
        logic.UseTunnelRL = Me.optTunnelRL.value

        ' Memanggil metode ApplyBasedOnOption dari class logic
        logic.ApplyBasedOnOption
    End If
    
    ' Kondisi untuk optBounding
   If optBounding.value Then

        ' Pilih apakah bounding box mencakup outline
        boxer.UseOutline = False

        ' Jalankan
        boxer.CreateBoundingBoxPerShape
    End If
    
CleanUp:
    ' Akhiri grup command meskipun ada kesalahan input
    ActiveDocument.EndCommandGroup
    Exit Sub

errHandler:
    MsgBox "Terjadi kesalahan: " & Err.Description, vbCritical
    Resume CleanUp
    
End Sub

' \\\\\\\\\\\\\\\
' ===============(Measure/Pengukuran)===============
Private Sub cmdMeasure_Click()
    If ActiveSelectionRange.Count = 0 Then
        MsgBox "Pilih satu shape terlebih dahulu!", vbExclamation
        Exit Sub
    End If

    Dim s As shape
    Set s = ActiveSelectionRange(1)

    Dim dimObj As clsAutoDimension
    Set dimObj = New clsAutoDimension
    dimObj.SetShape s

    ' Gambar dimensi horizontal (lebar)
    dimObj.DrawHorizontal

    ' Gambar dimensi vertikal (tinggi)
    dimObj.DrawVertical
End Sub

Private Sub SpinButton1_Change()
    txtKolongAB.value = spinKolAB.value
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
' ===============(Exit)===============
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGetLabels_Click()
    Unload Me
    frmLabelsBanner.Show VBMODELES
End Sub
