VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsQueLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Change()

Private WithEvents oLabel As Label
Attribute oLabel.VB_VarHelpID = -1

Public Enable As Boolean
Public EnE As String
Public EnO As String
Public Capp As String
Public Title As String
Public index As Long

Public Sub Init(ByVal lb As Label, ByVal cap As String)

    Set oLabel = lb
    Capp = cap
    index = 1
    
    poced
    
End Sub

Public Sub oLabel_Click()


    index = index + 1
    index = IIf(index > Len(Capp), 1, index)
    
    poced
    
    End Sub
    
Private Sub poced()

    
    Title = Mid$(Capp, index, 1)

    'oLabel.BackColor = IIf(Title = " ", vbGrayText, vbInfoBackground)
    oLabel.Caption = Title
    Select Case Title
    Case " "
        oLabel.Caption = "<>"
        EnE = " not "
        EnO = " = "
    Case "="
        EnE = " "
        EnO = " = "
    Case ">"
        EnO = " > "
    Case "<"
        EnO = " < "
    End Select

'    oLabel.BackColor = IIf(oLabel.BackColor = vbInfoBackground, vbGrayText, vbInfoBackground)
'    Enable = IIf(Enable, 0, 1)
'    oLabel.ControlTipText = IIf(Enable, "Equal", "NOT Equal")
'    EnE = IIf(Enable, " ", " not ")
    
    
End Sub

