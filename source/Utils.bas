Attribute VB_Name = "Utils"

Private InternalPrecision As Long
Private Value As Long
Private Unit As cdrUnit

Private Sub UpdatePrecision()
    If Precision < 0 Then
        InternalPrecision = GetUnitPrecision(Unit)
    Else
        InternalPrecision = Precision
    End If
End Sub


Private Function GetUnitPrecision(Unit As cdrUnit) As Long
'    Dim n As Long
'    Select Case Unit
'        Case cdrInch
'            n = 3
'        Case cdrPoint, cdrPixel
'            n = 0
'        Case Else
'            n = 2
'    End Select
    GetUnitPrecision = 3
End Function





