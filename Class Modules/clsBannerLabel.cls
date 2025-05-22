VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBannerLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsBannerLabel.cls
Option Explicit

Public CustomerName As String
Public Material As String
Public Finishing As String
Public Quantity As String
Public Address As String
Public ShowDate As Boolean
Public ShowAddress As Boolean
Public IsExpress As Boolean

Public Function GetLabelText(shapeWidth As Double, shapeHeight As Double) As String
    Dim result As String
    Dim currentDate As String
    Dim expressText As String
    Dim dateText As String
    Dim addressText As String

    currentDate = Format(Date, "dd/mm/yyyy")
    expressText = IIf(IsExpress, "- EXPRESS", "")
    dateText = IIf(ShowDate, "- @" & currentDate, "")
    addressText = IIf(ShowAddress, "- #" & Address, "")

    result = Round(shapeWidth, 2) & " x " & Round(shapeHeight, 2) & " (m) - " & _
             UCase(Material) & " " & expressText & " - " & UCase(CustomerName) & _
             " - " & UCase(Finishing) & " - " & Quantity & "X " & addressText & " " & dateText

    GetLabelText = result
End Function


