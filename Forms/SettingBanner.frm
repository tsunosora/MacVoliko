VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingBanner 
   Caption         =   "Banner Helper Setting"
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "SettingBanner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'--------------------------------------------------
' ============ Fungsi - Ketika UserForm di Buka ============
'--------------------------------------------------
Sub UserForm_Initialize() '------------------------------------------------ +++++ Fungsi UserForm
    
    'Warna Bacground label saat UserForm pertama kali di buka
    '----------------------------------------------------------------------------------------------------------------------------***
    Me.BackColor = RGB(0, 0, 94) 'Biru

    Call BacaFileTempat
End Sub
Private Sub CmdTutup_Click() '------------------------------------------------ +++++ Fungsi UserForm
    Unload Me
End Sub

'==================================================
' SUBrotin untuk Membaca file supaya nampak di TextBox
'==================================================
' ++ (Membaca file supaya berada di txtAlamat) ++
Private Sub BacaFileTempat() '------------------------------------------------ +++++ Fungsi UserForm
    Dim filePath As String
    Dim fileNum As Integer
    Dim fileContent As String
    
    ' Tentukan path file yang akan dibaca
    filePath = Environ("TEMP") & "\txtAlamat.txt" ' Ganti dengan path file yang sesuai
    
    ' Periksa apakah file ada
    If Dir(filePath) = "" Then
        
        Exit Sub
    End If
    
    ' Membuka file untuk dibaca
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Membaca seluruh isi file
    fileContent = Input$(LOF(fileNum), fileNum)
    
    ' Menutup file setelah selesai membaca
    Close fileNum
    
    ' Menghapus semua jenis spasi dan karakter kosong dari teks (termasuk spasi biasa, tab, dan spasi lainnya)
    fileContent = Replace(fileContent, " ", "") ' Menghapus semua spasi biasa
    fileContent = Replace(fileContent, vbTab, "") ' Menghapus semua tab
    fileContent = Replace(fileContent, vbCrLf, "") ' Menghapus semua karakter baris baru (jika ada)
    
    ' Menampilkan isi file ke dalam txtAlamat
    txtAlamat.text = fileContent
End Sub

'--------------------------------------------------
' ============ Fungsi - btnTambahBahan ============
'--------------------------------------------------
Private Sub btnTambahBahan_Click() '------------------------------------------------ +++++ Fungsi UserForm
    Call SerchFileB
End Sub

'==================================================
' SUBrotine Untuk mencari file #berdasarkan Nama file
'==================================================
' ++ (Mencari SvLstBn.txt dan Membukanya) ++
Sub SerchFileB()
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim folderPath As String
    Dim fileName As String
    Dim filePath As String

    ' Nama file yang ingin dicari
    fileName = "SvLstBn.txt"
    
    ' Folder tempat pencarian dimulai
    folderPath = "C:\Program Files\Corel" ' Ganti dengan folder yang kamu inginkan
    
    ' Membuat objek FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cek apakah folder ada dan dapat diakses
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder tidak ditemukan atau tidak dapat diakses!"
        Exit Sub
    End If
    
    ' Mendapatkan objek folder
    Set folder = fso.GetFolder(folderPath)
    
    ' Mulai pencarian file di folder dan subfolder
    Call CariFileDiSubfolder(folder, fileName, filePath)
    
    ' Jika file ditemukan, buka file dengan aplikasi default
    If filePath <> "" Then
        Shell "cmd /c start """" """ & filePath & """", vbHide
    End If
End Sub

'==================================================
' Lanjutan dari Subrotin SerchFileB()
' (SUBrotine Untuk mencari file #berdasarkan Nama file)
'==================================================
Private Sub CariFileDiSubfolder(ByVal folder As Object, ByVal fileName As String, ByRef filePath As String)
    Dim subfolder As Object
    Dim file As Object

    ' Mencari file di folder ini
    For Each file In folder.Files
        ' Membandingkan nama file dengan nama yang dicari
        If LCase(file.Name) = LCase(fileName) Then
            filePath = file.Path
            Exit Sub
        End If
    Next file
    
    ' Jika file tidak ditemukan di folder ini, periksa subfolder
    For Each subfolder In folder.SubFolders
        Call CariFileDiSubfolder(subfolder, fileName, filePath) ' Panggil rekursif untuk subfolder
    Next subfolder
End Sub

'--------------------------------------------------
' ============ Fungsi - CmdSimpan ============
'--------------------------------------------------
Private Sub cmdSimpan_Click()

    ' Panggil subroutine BuatAlamat dan kirimkan parameter fileName dan filePath
    BuatAlamat fileName, filePath
    
    ' menutup UserForm
    Unload Me
End Sub

'==================================================
' SUBrotin untuk Membuat file txtAlamat.txt
'==================================================
Private Sub BuatAlamat(ByVal fileName As String, ByVal filePath As String)
    Dim fso As Object
    Dim tempFile As Object
    Dim fileSavePath As String
    
    ' Tentukan path untuk menyimpan file sementara berdasarkan nama file
    fileSavePath = Environ("TEMP") & "\txtAlamat.txt"
    
    ' Membuat objek FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Membuat atau membuka file sementara untuk menulis data
    Set tempFile = fso.CreateTextFile(fileSavePath, True) ' True berarti akan menimpa file yang ada
    
    ' Menulis path dari txtPath ke dalam file
    tempFile.WriteLine txtAlamat.text ' Menulis nilai dari txtPath.Text
    
    ' Menutup file
    tempFile.Close
    
End Sub


'///////////////////////////////////////////////////////////////////////////////////////////////#Tidak di jalankan
'==================================================
' SUBroutine Untuk Mencari dan Menyalin 4 File .png
' dengan Memeriksa apakah file sudah ada di folder TEMP
'==================================================
Sub CopyMultiplePNGFilesToTemp()
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim folderPath As String
    Dim tempFolderPath As String
    Dim filePath As String
    Dim fileNames As Variant
    Dim i As Integer
    
    ' Daftar nama file .png yang ingin disalin
    fileNames = Array("Rset1.png", "Rset2.png", "Rset3.png", "Rset4.png")
    
    ' Tentukan folder awal untuk pencarian
    folderPath = "C:\Program Files\Corel" ' Ganti dengan folder yang sesuai
    
    ' Tentukan folder TEMP untuk menyalin file
    tempFolderPath = Environ("TEMP") ' Folder TEMP sistem
    
    ' Membuat objek FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cek apakah folder awal ada dan dapat diakses
    If Not fso.FolderExists(folderPath) Then
        Exit Sub
    End If
    
    ' Mendapatkan objek folder
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop untuk menyalin file yang ada di dalam daftar fileNames
    For i = LBound(fileNames) To UBound(fileNames)
        ' Pencarian dan penyalinan untuk setiap file dalam array fileNames
        If Not CariDanSalinFile(folder, fileNames(i), tempFolderPath) Then
        End If
    Next i
End Sub
' ========= Fungsi lamjutan dari SUBrotin CopyMultiplePNGFilesToTemp =========
'==================================================
' SUBroutine untuk Mencari dan Menyalin File
'==================================================
Function CariDanSalinFile(ByVal folder As Object, ByVal fileName As String, ByVal tempFolderPath As String) As Boolean
    Dim subfolder As Object
    Dim file As Object
    Dim filePath As String
    Dim tempFilePath As String
    
    ' Tentukan path lengkap file di folder TEMP
    tempFilePath = tempFolderPath & "\" & fileName
    
    ' Cek apakah file sudah ada di folder TEMP
    If FileExistsInTemp(tempFilePath) Then
        CariDanSalinFile = False
        Exit Function
    End If
    
    ' Loop untuk mencari file dalam folder ini
    For Each file In folder.Files
        If LCase(file.Name) = LCase(fileName) Then ' Memeriksa nama file
            filePath = file.Path
            ' Salin file ke folder TEMP
            FileCopy filePath, tempFilePath
            CariDanSalinFile = True ' Mengembalikan True setelah file disalin
            Exit Function ' Keluar setelah file ditemukan dan disalin
        End If
    Next file
    
    ' Jika file belum ditemukan, lanjutkan pencarian di subfolder
    For Each subfolder In folder.SubFolders
        If CariDanSalinFile(subfolder, fileName, tempFolderPath) Then
            CariDanSalinFile = True ' Mengembalikan True setelah file disalin
            Exit Function ' Keluar jika file ditemukan di subfolder
        End If
    Next subfolder
    
    CariDanSalinFile = False ' Mengembalikan False jika file tidak ditemukan
End Function
' ========= Fungsi lamjutan dari SUBrotin CopyMultiplePNGFilesToTemp =========
'==================================================
' Fungsi untuk Memeriksa Apakah File Sudah Ada di Folder TEMP
'==================================================
Function FileExistsInTemp(ByVal tempFilePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Mengecek apakah file sudah ada di folder TEMP
    If fso.FileExists(tempFilePath) Then
        FileExistsInTemp = True
    Else
        FileExistsInTemp = False
    End If
End Function





