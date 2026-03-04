Attribute VB_Name = "JobHunt"
'==============================================================================
' Module     : JobHunt
' Description: Main module for the JobHunt Outlook Add-in
'              Tracks job applications and automates job-hunting emails.
' Author     :
' Date       : 2026-03-04
'==============================================================================
Option Explicit

' ---- Constants ---------------------------------------------------------------
Public Const CATEGORY_JOB_APP   As String = "Job Application"
Public Const CATEGORY_RECRUITER As String = "Recruiter"
Public Const CATEGORY_FOLLOWUP  As String = "Follow-Up"
Public Const ADDIN_FOLDER       As String = "JobHunt"

' ---- Public Entry Points -----------------------------------------------------

'------------------------------------------------------------------------------
' Marks the currently selected email as a job application and categories it.
'------------------------------------------------------------------------------
Public Sub TagAsJobApplication()
    Dim oMail As Outlook.MailItem
    Set oMail = GetSelectedMail()
    If oMail Is Nothing Then Exit Sub

    oMail.Categories = CATEGORY_JOB_APP
    oMail.Save
    MsgBox "Email tagged as '" & CATEGORY_JOB_APP & "'.", vbInformation, "JobHunt"
End Sub

'------------------------------------------------------------------------------
' Marks the currently selected email as a recruiter contact.
'------------------------------------------------------------------------------
Public Sub TagAsRecruiter()
    Dim oMail As Outlook.MailItem
    Set oMail = GetSelectedMail()
    If oMail Is Nothing Then Exit Sub

    oMail.Categories = CATEGORY_RECRUITER
    oMail.Save
    MsgBox "Email tagged as '" & CATEGORY_RECRUITER & "'.", vbInformation, "JobHunt"
End Sub

'------------------------------------------------------------------------------
' Marks the currently selected email as needing a follow-up.
'------------------------------------------------------------------------------
Public Sub TagAsFollowUp()
    Dim oMail As Outlook.MailItem
    Set oMail = GetSelectedMail()
    If oMail Is Nothing Then Exit Sub

    oMail.Categories = CATEGORY_FOLLOWUP
    oMail.FlagRequest = "Follow up"
    oMail.FlagDueBy = Now + 3  ' 3-day reminder
    oMail.Save
    MsgBox "Follow-up flag set (due in 3 days).", vbInformation, "JobHunt"
End Sub

'------------------------------------------------------------------------------
' Shows a summary of all job-related emails in a message box.
'------------------------------------------------------------------------------
Public Sub ShowJobSummary()
    Dim oNS     As Outlook.NameSpace
    Dim oItems  As Outlook.Items
    Dim oMail   As Object
    Dim nApp    As Long
    Dim nRec    As Long
    Dim nFup    As Long

    Set oNS    = Application.GetNamespace("MAPI")
    Set oItems = oNS.GetDefaultFolder(olFolderInbox).Items

    Dim i As Long
    For i = 1 To oItems.Count
        Set oMail = oItems(i)
        If oMail.Class = olMail Then
            Dim cats As String
            cats = oMail.Categories
            If InStr(cats, CATEGORY_JOB_APP) > 0   Then nApp = nApp + 1
            If InStr(cats, CATEGORY_RECRUITER) > 0 Then nRec = nRec + 1
            If InStr(cats, CATEGORY_FOLLOWUP) > 0  Then nFup = nFup + 1
        End If
    Next i

    MsgBox "JobHunt Summary (Inbox)" & vbNewLine & vbNewLine & _
           "Applications : " & nApp & vbNewLine & _
           "Recruiters   : " & nRec & vbNewLine & _
           "Follow-Ups   : " & nFup, _
           vbInformation, "JobHunt Summary"
End Sub

' ---- Job Application Dialog -------------------------------------------------

'------------------------------------------------------------------------------
' Opens the Job Application management dialog for the currently selected or
' open mail.  Bound to the right-click context menu item installed at startup.
'------------------------------------------------------------------------------
Public Sub ShowJobApplicationDialog()
    Dim oMail As Outlook.MailItem
    Set oMail = GetCurrentMailItem()

    If oMail Is Nothing Then
        MsgBox "Bitte zuerst eine E-Mail auswaehlen oder oeffnen.", _
               vbExclamation, "JobHunt"
        Exit Sub
    End If

    ' Ensure settings file exists before showing the dialog
    Settings.EnsureSettingsExist

    frmJobApplication.Show vbModeless
End Sub

' ---- Context Menu ------------------------------------------------------------

Private Const CTX_BTN_CAPTION As String = "Job Application..."

'------------------------------------------------------------------------------
' Installs a "Job Application..." button in the Outlook mail item context menu.
' Called from Application_Startup in ThisOutlookSession.
'------------------------------------------------------------------------------
Public Sub InstallContextMenu()
    RemoveContextMenu   ' avoid duplicates

    On Error Resume Next
    Dim oCB  As Office.CommandBar
    Dim oBtn As Office.CommandBarButton

    ' Outlook uses different CommandBar names; try common ones
    Dim aNames(2) As String
    aNames(0) = "Context Menu"
    aNames(1) = "Message"
    aNames(2) = "Item List"

    Dim i As Integer
    For i = 0 To 2
        Set oCB = Application.ActiveExplorer.CommandBars(aNames(i))
        If Not oCB Is Nothing Then
            Set oBtn = oCB.Controls.Add(msoControlButton, , , , True)
            oBtn.Caption    = CTX_BTN_CAPTION
            oBtn.OnAction   = "ShowJobApplicationDialog"
            oBtn.Style      = msoButtonIconAndCaption
            oBtn.BeginGroup = True
        End If
        Set oCB = Nothing
    Next i
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Removes the "Job Application..." context menu button.
' Called from Application_Quit in ThisOutlookSession.
'------------------------------------------------------------------------------
Public Sub RemoveContextMenu()
    On Error Resume Next
    Dim oCB  As Office.CommandBar
    Dim oCtrl As Office.CommandBarControl

    For Each oCB In Application.ActiveExplorer.CommandBars
        For Each oCtrl In oCB.Controls
            If oCtrl.Caption = CTX_BTN_CAPTION Then
                oCtrl.Delete
            End If
        Next oCtrl
    Next oCB
    On Error GoTo 0
End Sub

' ---- Helper Functions --------------------------------------------------------

'------------------------------------------------------------------------------
' Returns the current MailItem from the open Inspector or Explorer selection.
' Returns Nothing if no mail is available.
'------------------------------------------------------------------------------
Public Function GetCurrentMailItem() As Outlook.MailItem
    ' 1. Open Inspector
    On Error Resume Next
    Dim oInsp As Outlook.Inspector
    Set oInsp = Application.ActiveInspector
    If Not oInsp Is Nothing Then
        If oInsp.CurrentItem.Class = olMail Then
            Set GetCurrentMailItem = oInsp.CurrentItem
            Exit Function
        End If
    End If
    On Error GoTo 0

    ' 2. Explorer selection
    On Error Resume Next
    Dim oSel As Outlook.Selection
    Set oSel = Application.ActiveExplorer.Selection
    If oSel.Count > 0 Then
        If oSel(1).Class = olMail Then
            Set GetCurrentMailItem = oSel(1)
        End If
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Returns the first selected MailItem, or Nothing when none is selected.
'------------------------------------------------------------------------------
Private Function GetSelectedMail() As Outlook.MailItem
    Dim oSel As Outlook.Selection
    Set oSel = Application.ActiveExplorer.Selection

    If oSel.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "JobHunt"
        Exit Function
    End If

    If oSel(1).Class <> olMail Then
        MsgBox "The selected item is not an email.", vbExclamation, "JobHunt"
        Exit Function
    End If

    Set GetSelectedMail = oSel(1)
End Function
