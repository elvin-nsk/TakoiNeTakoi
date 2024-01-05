VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public ColorTypeName As String
Public ColorHSV_H As Long
Public ColorHSV_S As Long
Public ColorHSV_V As Long
Public ColorHEX As String
Public ColorRGB_R As Long
Public ColorRGB_G As Long
Public ColorRGB_B As Long



Public Function getColorTypeName(colortype As Integer) As String
    Dim n As String
    
    
    Select Case colortype
    Case 1
        S = "cdrColorPantone 1   Specifies PANTONE"
    Case 2
        S = "cdrColorCMYK    2   Specifies CMYK"
    Case 4
        S = "cdrColorCMY     4   Specifies CMY"
    Case 5
        S = "cdrColorRGB     5   Specifies RGB"
    Case 6
        S = "cdrColorHSB     6   Specifies HSB"
    Case 7
        S = "cdrColorHLS     7   Specifies HLS"
    Case 8
        S = "cdrColorBlackAndWhite    8   Specifies black-and-white"
    Case 9
        S = "cdrColorGray    9   Specifies grayscale"
    Case 11
        S = "cdrColorYIQ    11  Specifies YIQ"
    Case 12
        S = "cdrColorLab    12  Specifies LAB"
    Case 14
        S = "cdrColorPantoneHex      14  Specifies PANTONE Hex"
    Case 20
        S = "cdrColorRegistration    20  Specifies registration"
    Case 22
        S = "cdrColorUserInk 22  Specifies multi-channel"
    Case 25
        S = "cdrColorSpot    25  Specifies user ink"
    Case 26
        S = "cdrColorMultiChannel    26  Specifies mixed"
    Case 99
        S = "cdrColorMixed   99  Specifies spot"
    End Select

    getColorTypeName = S
    ColorTypeName = S
    
End Function


Private Function max(x, y, z As Variant) As Variant
    Dim tmax As Variant
    tmax = IIf(x > y, x, y)
    tmax = IIf(tmax > z, tmax, z)
    max = tmax
End Function

Private Function min(x, y, z As Variant) As Variant
    Dim tmix As Variant
    tmin = IIf(x < y, x, y)
    tmin = IIf(tmin < z, tmin, z)
    min = tmin
End Function


Public Sub GetHSV_From_RGB(R As Long, G As Long, B As Long)
    Dim ReturnValue As Long
 
    Dim H, S, V As Long
 
    Dim TempH, TempS, TempV As Double
 
    Dim TempR, TempG, TempB As Double
 
    Dim MinRGBInteger, MaxRGBInteger As Long
    Dim MinRGBDouble, MaxRGBDouble As Double
 
    TempR = R / 255
    TempG = G / 255
    TempB = B / 255
 
    MinRGBInteger = min(R, G, B)
    MaxRGBInteger = max(R, G, B)
 
    MinRGBDouble = MinRGBInteger / 255
    MaxRGBDouble = MaxRGBInteger / 255
 
    V = MaxRGBInteger
 
    TempV = V / 255
 
    If (V <> 0) Then
        TempS = (TempV - MinRGBDouble) / TempV
    Else
        TempS = 0
    End If
 
    If (MinRGBInteger <> MaxRGBInteger) Then
        If (V = R) Then
            TempH = ((60 * (TempG - TempB)) / (TempV - MinRGBDouble))
        End If
 
        If (V = G) Then
            TempH = 120 + (60 * (TempB - TempR) / (TempV - MinRGBDouble))
        End If
 
        If (V = B) Then
            TempH = 240 + (60 * (TempR - TempG) / (TempV - MinRGBDouble))
        End If
 
        If (TempH < 0) Then
            TempH = TempH + 360
        End If
    Else
        TempH = 0
    End If
 
    H = TempH / 2
    S = 255 * TempS
    
    ColorHSV_H = H
    ColorHSV_S = S
    ColorHSV_V = V
 
End Sub


Public Function GetHex_From_RGB(Red As Long, Green As Long, Blue As Long) As String

    ColorHEX = "#" & VBA.Right$("00" & VBA.Hex(Red), 2) & _
        VBA.Right$("00" & VBA.Hex(Green), 2) & VBA.Right$("00" & VBA.Hex(Blue), 2)
    
    GetHex_From_RGB = ColorHEX

End Function

Public Function GetRGB_From_Hex(hexColor As String, Optional RGB As String = "") As String

    hexColor = VBA.Replace(hexColor, "#", "")
    hexColor = VBA.Right$("000000" & hexColor, 6)


    ColorRGB_R = VBA.Val("&H" & VBA.Mid(hexColor, 5, 2))
    ColorRGB_G = VBA.Val("&H" & VBA.Mid(hexColor, 3, 2))
    ColorRGB_B = VBA.Val("&H" & VBA.Mid(hexColor, 1, 2))


Select Case RGB

    Case "B"
        GetRGB_From_Hex = VBA.Val("&H" & VBA.Mid(hexColor, 5, 2))

    Case "G"
        GetRGB_From_Hex = VBA.Val("&H" & VBA.Mid(hexColor, 3, 2))

    Case "R"
        GetRGB_From_Hex = VBA.Val("&H" & VBA.Mid(hexColor, 1, 2))

End Select

End Function

Function GetLong_From_RGB(Red As Long, Green As Long, Blue As Long) As Long

    GetLong_From_RGB = RGB(Red, Green, Blue)
End Function


Function GetRGB_From_Long(longColor As Long, Optional RGB As String = "") As Long

    ColorRGB_R = (longColor Mod 256)
    ColorRGB_G = (longColor \ 256) Mod 256
    ColorRGB_B = (longColor \ 65536) Mod 256

Select Case RGB
    Case "R"
        GetRGB_From_Long = (longColor Mod 256)

    Case "G"
        GetRGB_From_Long = (longColor \ 256) Mod 256

    Case "B"
        GetRGB_From_Long = (longColor \ 65536) Mod 256

End Select

End Function
