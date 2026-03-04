VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : clsOptSink
' Description: WithEvents sink for dynamically-created OptionButtons.
'              Used by frmJobApplication to handle option-button clicks
'              without design-time controls.
'==============================================================================
Option Explicit

Public WithEvents Opt As MSForms.OptionButton
Public Form As Object   ' reference back to frmJobApplication

Private Sub Opt_Click()
    Form.OnOptionClick Opt.Name
End Sub
