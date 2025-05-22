VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWarnaPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === clsWarnaPicker ===
Private cColor As Color
Private cWarnaTerpilih As Color

' Fungsi untuk menampilkan dialog warna dan menyimpan hasilnya
Public Function PilihWarna() As Boolean
    Set cColor = New Color
    
    If cColor.UserAssignEx = True Then
        Set cWarnaTerpilih = cColor
        PilihWarna = True
    Else
        Set cWarnaTerpilih = CreateCMYKColor(0, 0, 0, 0) ' fallback warna default
        PilihWarna = False
    End If
End Function

' Ambil warna yang dipilih
Public Property Get WarnaTerpilih() As Color
    Set WarnaTerpilih = cWarnaTerpilih
End Property

' Ambil nilai CMYK dari warna
Public Property Get CMYKCyan() As Double
    CMYKCyan = cWarnaTerpilih.CMYKCyan
End Property

Public Property Get CMYKMagenta() As Double
    CMYKMagenta = cWarnaTerpilih.CMYKMagenta
End Property

Public Property Get CMYKYellow() As Double
    CMYKYellow = cWarnaTerpilih.CMYKYellow
End Property

Public Property Get CMYKBlack() As Double
    CMYKBlack = cWarnaTerpilih.CMYKBlack
End Property

' Konversi ke warna RGB (untuk ditampilkan di UserForm)
Public Function GetWarnaRGB() As Long
    Dim c As New Color
    c.CMYKAssign cWarnaTerpilih.CMYKCyan, _
                 cWarnaTerpilih.CMYKMagenta, _
                 cWarnaTerpilih.CMYKYellow, _
                 cWarnaTerpilih.CMYKBlack
    c.ConvertToRGB
    GetWarnaRGB = RGB(c.RGBRed, c.RGBGreen, c.RGBBlue)
End Function


