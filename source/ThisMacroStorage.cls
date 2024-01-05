VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisMacroStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Page As Page
Public WithEvents Doc As Document
Attribute Doc.VB_VarHelpID = -1
Public EventNumber As Long
Public EventNumberString As String
Public eventString As String
Public StopEvents As Boolean
Public ReadyForWorking As Boolean

Private Sub Doc_Close()
    eventString = "Doc_Close" & Doc.Name & " "
    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber
    
    Set Doc = Nothing
End Sub

Private Sub Doc_Open()
    eventString = "Doc_Open" & Doc.Name & " "
    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber

End Sub

Private Sub Doc_QueryClose(Cancel As Boolean)
    
    eventString = "Doc_QueryClose" & Doc.Name & " "
    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber
    
'    Set Doc = Nothing
End Sub


Private Sub Doc_SelectionChange()
    eventString = "Doc_SelectionChange " & Doc.Name & " "
    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber
    
End Sub

Private Sub Doc_ShapeChange(ByVal Shape As Shape, ByVal Scope As cdrShapeChangeScope)
    eventString = "Doc_ShapeChange " & Doc.Name & " "
    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber

End Sub


Private Sub GlobalMacroStorage_OnApplicationEvent(ByVal EventName As String, Parameters() As Variant)

    
    'If Not ReadyForWorking Then Exit Sub


    If Application.ActiveDocument Is Nothing Then
        eventString = "-- OnApplicationEvent - " & EventName
    Else
        If Doc Is Nothing Then
            Set Doc = Application.ActiveDocument
        Else
            If Doc.Name <> Application.ActiveDocument Then
                Set Doc = Application.ActiveDocument
            End If
        End If
        eventString = "---OnApplicationEvent - " & EventName & " - " & Doc.Name & " "
    End If


    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber
    
    Select Case EventName
    
    Case "DocUnitChange", "WorkspaceChanged"
        Dim R As Variant
        R = Parameters()
        
        frmCloneAndRotate.queUnit.Init Application.ActiveDocument.Rulers.HUnits, Application.ActiveDocument.WorldScale
         
    
    End Select


End Sub

Private Sub GlobalMacroStorage_SelectionChange()

    If Application.ActiveDocument Is Nothing Then
        eventString = "MacroStorage_SelectionChange NOTHING"
    Else
        If Doc Is Nothing Then
            Set Doc = Application.ActiveDocument
        Else
            If Doc.Name <> Application.ActiveDocument Then
                Set Doc = Application.ActiveDocument
            End If
        End If
        eventString = "MacroStorage_SelectionChange " & Doc.Name & " "
    End If


    EventNumber = EventNumber + 1
    AddtoList eventString, EventNumber
    
    

End Sub



Sub AddtoList(event_string, event_number)
    
    On Error Resume Next

    'EventNumberString = Format$(event_number, "0." & String$(InternalPrecision, "0"))
    'frmCloneAndRotate.lstMessage.AddItem event_string & EventNumberString
    'frmCloneAndRotate.lstMessage.ListIndex = frmCloneAndRotate.lstMessage.ListCount - 1

End Sub
