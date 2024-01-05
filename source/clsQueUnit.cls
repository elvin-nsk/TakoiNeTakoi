VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WorldScale As Double
Public Precision As Long


'================= Private Data =================
Private InternalPrecision As Long
Private Value As Long
Public Unit As cdrUnit
Public UnitName As String
Public UnitShortName As String


Private Sub Class_Initialize()
    Unit = cdrInch
    Value = 0
    WorldScale = 1
End Sub


Public Sub InitOld()
    WorldScale = Application.ActiveDocument.WorldScale
    Unit = cdrInch
    Precision = -1
    UpdatePrecision
End Sub

Public Sub Init(ByVal NewUnit As cdrUnit, ByVal NewWorldScale As Double)
    WorldScale = NewWorldScale
    Unit = NewUnit
    Precision = GetUnitPrecision(NewUnit)
    setUnitName (Unit)
    UpdatePrecision
End Sub



Public Sub OnUnitChange(ByVal NewUnit As cdrUnit)
    Unit = NewUnit
    setUnitName (Unit)
    UpdatePrecision
End Sub

Public Sub OnWorldScaleChange(ByVal NewWorldScale As Double)
    WorldScale = NewWorldScale
    UpdatePrecision
End Sub



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
    GetUnitPrecision = 4
End Function


Public Function FromUnit(ByVal V As Double) As Double
    FromUnit = V / WorldScale * GetUnitFactor(Unit)
End Function

Public Function ToUnit(ByVal V As Double) As Double
    ToUnit = V / GetUnitFactor(Unit) * WorldScale
End Function

Public Function FormatValue(ByVal ln As Double) As String
    Dim V As Double
    Dim d As Double
    Dim sTrail As String, S As String
    Dim n As Long
    
    d = 10 ^ InternalPrecision
    V = ToUnit(ln)
    V = Fix(V * d + 0.5 * Sgn(V)) / d
    If InternalPrecision > 0 Then
        sTrail = Format$(V, "0." & String$(InternalPrecision, "0"))
        For n = Len(sTrail) To 1 Step -1
            S = Mid$(sTrail, n, 1)
            If S <> "0" Then
                If S > "0" And S <= "9" Then
                    sTrail = Mid$(sTrail, n + 1)
                Else
                    sTrail = "." & Mid$(sTrail, n + 1)
                End If
                Exit For
            End If
        Next n
    Else
        sTrail = ""
    End If
    sTrail = Trim$(Str$(V)) & sTrail
    If Left$(sTrail, 1) = "." Then
        sTrail = "0" & sTrail
    ElseIf Left$(sTrail, 2) = "-." Then
        sTrail = "-0" & Mid$(sTrail, 2)
    End If
    FormatValue = sTrail
End Function


Public Function GetUnitFactor(Unit As cdrUnit) As Double
    Dim dFactor As Double
    Select Case Unit
        Case cdrTenthMicron
            dFactor = 1
        Case cdrInch
            dFactor = 254000
        Case cdrCentimeter
            dFactor = 100000
        Case cdrMillimeter
            dFactor = 10000
        Case cdrFoot
            dFactor = 3048000
        Case cdrPixel
            dFactor = 846.6666395733
        Case cdrMile
            dFactor = 16093440000#
        Case cdrMeter
            dFactor = 10000000
        Case cdrKilometer
            dFactor = 10000000000#
        Case cdrDidots
            dFactor = 3759.2
        Case cdrAgate
            dFactor = 18142.85714286
        Case cdrYard
            dFactor = 9144000
        Case cdrPica
            dFactor = 42333.33333333
        Case cdrCicero
            dFactor = 45118.7
        Case cdrPoint
            dFactor = 3527.77777777778
        Case cdrUnitH, cdrUnitQ
            dFactor = 2500
    End Select
    GetUnitFactor = dFactor
End Function

Public Function GetUnitName(Unit As cdrUnit) As String
    Dim dName As String
    Select Case Unit
        Case cdrTenthMicron
            dName = "TenthMicron"
        Case cdrInch
            dName = "Inch"
        Case cdrCentimeter
            dName = "Centimeter"
        Case cdrMillimeter
            dName = "Millimeter"
        Case cdrFoot
            dName = "Foot"
        Case cdrPixel
            dName = "Pixel"
        Case cdrMile
            dName = "Mile"
        Case cdrMeter
            dName = "Meter"
        Case cdrKilometer
            dName = "Kilometer"
        Case cdrDidots
            dName = "Didots"
        Case cdrAgate
            dName = "Agate"
        Case cdrYard
            dName = "Yard"
        Case cdrPica
            dName = "Pica"
        Case cdrCicero
            dName = "Cicero"
        Case cdrPoint
            dName = "Point"
        Case cdrUnitH
            dName = "H"
        Case cdrUnitQ
            dName = "Q"
    End Select
    GetUnitName = dName
End Function



Public Sub setUnitName(Unit As cdrUnit)

Dim dName As String
Dim sName As String

    Select Case Unit
        Case cdrTenthMicron
            dName = "TenthMicron"
        Case cdrInch
            dName = "Inch"
            sName = " in "
        Case cdrCentimeter
            dName = "Centimeter"
        Case cdrMillimeter
            dName = "Millimeter"
            sName = " mm "
        Case cdrFoot
            dName = "Foot"
        Case cdrPixel
            dName = "Pixel"
        Case cdrMile
            dName = "Mile"
        Case cdrMeter
            dName = "Meter"
        Case cdrKilometer
            dName = "Kilometer"
        Case cdrDidots
            dName = "Didots"
        Case cdrAgate
            dName = "Agate"
        Case cdrYard
            dName = "Yard"
        Case cdrPica
            dName = "Pica"
        Case cdrCicero
            dName = "Cicero"
        Case cdrPoint
            dName = "Point"
        Case cdrUnitH
            dName = "H"
        Case cdrUnitQ
            dName = "Q"
    End Select
    
    UnitName = dName
    UnitShortName = sName
    
End Sub


