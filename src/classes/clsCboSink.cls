VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCboSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : clsCboSink
' Description: WithEvents sink for dynamically-created ComboBoxes.
'              Used by frmJobApplication to handle combo-change events
'              without design-time controls.
'==============================================================================
Option Explicit

Public WithEvents Cbo As MSForms.ComboBox
Public Form As Object   ' reference back to frmJobApplication

Private Sub Cbo_Change()
    Form.OnComboChange Cbo.Name
End Sub
