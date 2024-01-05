VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCloneAndRotate 
   Caption         =   "Takoi ili Ne Takoi"
   ClientHeight    =   3165
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8730.001
   OleObjectBlob   =   "frmCloneAndRotate.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCloneAndRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public queShapeType As cdrShapeType
Public queUnit As clsQueUnit
Public ActiveSelectionRange_Item_1 As Shape
Public OutlineColorRGB As Long
Public FillUniformColorRGB As Long
Public OutlineWidth As Double

Dim WithEvents cShapeType As clsQueLabel
Attribute cShapeType.VB_VarHelpID = -1
Dim WithEvents cFillColor As clsQueLabel
Attribute cFillColor.VB_VarHelpID = -1
Dim WithEvents cOutlineColor As clsQueLabel
Attribute cOutlineColor.VB_VarHelpID = -1
Dim WithEvents cOutlineWidth As clsQueLabel
Attribute cOutlineWidth.VB_VarHelpID = -1



Private Sub UserForm_Initialize()
    ReadyForWorking = False
    Set queUnit = New clsQueUnit

    Set cShapeType = New clsQueLabel
    Set cFillColor = New clsQueLabel
    Set cOutlineColor = New clsQueLabel
    Set cOutlineWidth = New clsQueLabel
    
    cShapeType.Init Label23, "= "
    cFillColor.Init Label24, "= "
    cOutlineWidth.Init Label25, "=>< "
    cOutlineColor.Init Label26, "= "

    ReadyForWorking = True
End Sub

Private Sub lbFixAndFind_Click()

    lbFix_Click
    lbFind_Click
End Sub

Private Sub lbFind_Click()

    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then Exit Sub


    Dim findString As String
    Dim colorRGB As Long
    
    Dim R As Long
    Dim G As Long
    Dim B As Long
        
    Dim EnE As String
    
    
    If chOutlineColor And Label19.Visible Then
    
        colorRGB = OutlineColorRGB
        R = colorRGB \ 256 ^ 0 And 255
        G = colorRGB \ 256 ^ 1 And 255
        B = colorRGB \ 256 ^ 2 And 255
        findString = cOutlineColor.EnE & " @outline.color.rgb[.r=" & R & " and .g=" & G & " and .b=" & B & "] "

    End If
    
    If chOutlineColor And Not Label19.Visible Then
        If findString <> "" Then findString = findString & " and "
        findString = cOutlineColor.EnE & " @outline.type = 'none' "
    End If
    
    
    
    If chFillColor And Label18.Visible Then
        If findString <> "" Then findString = findString & " and "
    
        colorRGB = FillUniformColorRGB
        R = colorRGB \ 256 ^ 0 And 255
        G = colorRGB \ 256 ^ 1 And 255
        B = colorRGB \ 256 ^ 2 And 255
        findString = cFillColor.EnE & " @fill.color.rgb[.r=" & R & " and .g=" & G & " and .b=" & B & "] "

    End If
    
    
    If chFillColor And Not Label18.Visible Then
        If findString <> "" Then findString = findString & " and "
        findString = cFillColor.EnE & " @fill.type = 'none' "
    End If
    
    
    
    
    
    If chOutlineWidth Then
        If findString <> "" Then findString = findString & " and "
       
        findString = findString & cOutlineWidth.EnE & " @outline.width " & cOutlineWidth.EnO & " {" & queUnit.ToUnit(OutlineWidth) & queUnit.UnitShortName & "}"
    End If
    
    
    
    If chShape Then
        Select Case queShapeType
        Case 1
           If findString <> "" Then findString = findString & " and "
           findString = findString & cShapeType.EnE & " @type = 'rectangle' "
        Case 2
           If findString <> "" Then findString = findString & " and "
           findString = findString & cShapeType.EnE & " @type = 'ellipse' "
        Case 3
           If findString <> "" Then findString = findString & " and "
           findString = findString & cShapeType.EnE & " @type = 'curve' "
        Case 4
           If findString <> "" Then findString = findString & " and "
           findString = findString & cShapeType.EnE & " @type = 'polygon' "
        Case 6
           If findString <> "" Then findString = findString & " and "
           findString = findString & cShapeType.EnE & " @type = 'text:artistic' "
        End Select
    End If
    
    
    If findString = "" Then Exit Sub

    If (opPage.Value) Then
        ActivePage.Shapes.FindShapes(Query:=findString).CreateSelection
    End If

    If (opLayer.Value) Then
        ActivePage.ActiveLayer.Shapes.FindShapes(Query:=findString).CreateSelection
    End If


    'If (opGroup.Value And (Not ActiveSelectionRange.Item(1).ParentGroup Is Nothing)) Then
    '    ActiveSelectionRange.Item(1).ParentGroup.FindShapes(Query:=findString).CreateSelection
    'End If


    If (opSelection.Value) Then
        ActiveSelectionRange.Shapes.FindShapes(Query:=findString).CreateSelection
    End If



End Sub



Private Sub cmdCancel_Click()

    End
End Sub

Private Sub cmStop_Click()
    StopEvents = StopEvents <> 0
    eventsenable = StopEvents
    
End Sub



Private Sub lbFix_Click()
    
    
    Label14.Caption = ""
    Label15.Caption = ""
    
    
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveSelectionRange.Count = 0 Then Exit Sub
    
    ActiveDocument.SaveSettings "Clone" ' Save the current document settings first
    ActiveDocument.Unit = cdrTenthMicron
    
    Dim ooo As Shape
    Dim www As Double
    
    Set ActiveSelectionRange_Item_1 = ActiveSelectionRange.Item(1)
    
    queShapeType = ActiveSelectionRange.Item(1).Type
       
    
    Dim DocUnit As cdrUnit
    DocUnit = ActiveDocument.Rulers.HUnits
    
    
    queShapeType = ActiveSelectionRange.Item(1).Type
    Label14.Caption = getShapeTypeName(queShapeType)
    
    
    Select Case queShapeType
    Case 1, 2, 3, 4, 6
        OutlineWidth = ActiveSelectionRange.Item(1).Outline.Width
        
        queUnit.Init Application.ActiveDocument.Rulers.HUnits, Application.ActiveDocument.WorldScale
        
        Label15.Caption = queUnit.FormatValue(OutlineWidth) & " " & queUnit.GetUnitName(queUnit.Unit)
        
        
        Dim clorType As String
        Dim clors() As String
        Dim i As Integer
        Dim n As Integer
        
        
        FillUniformColorRGB = ActiveSelectionRange.Item(1).Fill.UniformColor.RGBValue
        Label18.BackColor = ActiveSelectionRange.Item(1).Fill.UniformColor.RGBValue
        Label18.Visible = IIf(ActiveSelectionRange.Item(1).Fill.Type = cdrNoFill, False, True)
        Label20.Caption = ActiveSelectionRange.Item(1).Fill.UniformColor.HexValue
        
        
        clorType = ActiveSelectionRange.Item(1).Fill.UniformColor.ToString
        clors = Split(clorType, ",")
        Label27.Caption = clors(0)
    
        n = 0
        For Each a In clors
            n = n + 1
        Next a
    
        For i = 2 To n - 3
            Label27.Caption = Label27.Caption & ", " & clors(i)
        Next i
        
        
        
        
        OutlineColorRGB = ActiveSelectionRange.Item(1).Outline.Color.RGBValue
        Label19.BackColor = ActiveSelectionRange.Item(1).Outline.Color.RGBValue
        Label19.Visible = IIf(ActiveSelectionRange.Item(1).Outline.Type = cdrNoOutline, False, True)
        Label21.Caption = ActiveSelectionRange.Item(1).Outline.Color.HexValue
        
        
        clorType = ActiveSelectionRange.Item(1).Outline.Color.ToString
        clors = Split(clorType, ",")
        Label22.Caption = clors(0)
    
        n = 0
        For Each a In clors
            n = n + 1
        Next a
    
        For i = 2 To n - 3
            Label22.Caption = Label22.Caption & ", " & clors(i)
        Next i
        
        
        
    End Select
    
    ActiveDocument.RestoreSettings "Clone"
    End Sub
    
    
Function getShapeTypeName(ttt As cdrShapeType) As String
    Dim sss As String
    
    Select Case ttt
    Case 0
        sss = "Specifies no shape"
    Case 1
        sss = "Specifies rectangle"
    Case 2
        sss = "Specifies ellipse"
    Case 3
        sss = "Specifies curve"
    Case 4
        sss = "Specifies polygon"
    Case 5
        sss = "Specifies bitmap"
    Case 6
        sss = "Specifies text"
    Case 7
        sss = "Specifies group"
    Case 8
        sss = "Specifies selection"
    Case 9
        sss = "Specifies guideline"
    Case 10
        sss = "Specifies blend group"
    Case 11
        sss = "Specifies extrude group"
    Case 12
        sss = "Specifies OLE object"
    Case 13
        sss = "Specifies contour group"
    Case 14
        sss = "Specifies linear dimension"
    Case 15
        sss = "Specifies bevel group"
    Case 16
        sss = "Specifies drop-shadow group"
    Case 17
        sss = "Specifies 3D object"
    Case 18
        sss = "Specifies artistic-media group"
    Case 19
        sss = "Specifies connector"
    Case 20
        sss = "Specifies mesh fill"
    Case 21
        sss = "Specifies custom shape"
    Case 22
        sss = "Specifies custom-effect group"
    Case 23
        sss = "Specifies symbol"
    Case 24
        sss = "Specifies HTML form object"
    Case 25
        sss = "Specifies HTML active object"
    Case 26
        sss = "Specifies Perfect Shape"
    Case 27
        sss = "Specifies EPS"
    End Select
    
    
    getShapeTypeName = sss

End Function
    


Private Sub lbClock_Click()
    
    FindMostContours

End Sub


Sub FindMostContours()
    Dim p As Page
    Dim pp As Shape
    Dim srContours As ShapeRange

    Set p = ActivePage
        Set srContours = p.Shapes.FindShapes(Query:="@type='rectangle'")
        If srContours.Shapes.Count > 0 Then
        
            For Each pp In srContours
                pp.Rotate 45
            Next pp
        End If
End Sub

Private Sub lbSelect_Click()

    'frmCloneAndRotate.MultiPage1.Value = 1

End Sub
Private Sub Label13_Click()

    'frmCloneAndRotate.MultiPage1.Value = 0
End Sub

