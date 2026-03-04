VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBtnSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : clsBtnSink
' Description: WithEvents sink for dynamically-created CommandButtons.
'              Used by frmJobApplication to handle button clicks without
'              requiring design-time controls (and therefore no .frx file
'              and no VBE/Application.VBE access).
'==============================================================================
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Public Form As Object   ' reference back to frmJobApplication

Private Sub Btn_Click()
    Form.OnButtonClick Btn.Name
End Sub
