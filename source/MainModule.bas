Attribute VB_Name = "MainModule"
Option Explicit



'----------------------------
' The main macro entry point
'----------------------------
Sub StartTakoiNeTakoi()

    Set ThisMacroStorage.Doc = ActiveDocument
    ThisMacroStorage.EventNumber = 0

    frmCloneAndRotate.Show
End Sub
