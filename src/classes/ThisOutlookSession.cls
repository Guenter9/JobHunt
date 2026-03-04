VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : ThisOutlookSession
' Description: Outlook application-level event hooks for the JobHunt Add-in.
'              Paste this code into the real ThisOutlookSession class module
'              inside the Outlook VBA IDE (Tools > Macros > Visual Basic Editor).
'==============================================================================
Option Explicit

Private WithEvents m_Inbox As Outlook.Items

'------------------------------------------------------------------------------
' Fired when Outlook starts.
' - Ensures settings file exists (creates defaults if missing)
' - Hooks into the Inbox Items collection for auto-tagging
' - Installs the right-click context menu button
'------------------------------------------------------------------------------
Private Sub Application_Startup()
    ' Ensure the %APPDATA%\JobHunt\settings.ini exists
    Settings.EnsureSettingsExist

    ' Hook into Inbox for auto-tagging
    Dim oNS As Outlook.NameSpace
    Set oNS     = Application.GetNamespace("MAPI")
    Set m_Inbox = oNS.GetDefaultFolder(olFolderInbox).Items

    ' Install context menu item
    JobHunt.InstallContextMenu

    Debug.Print "JobHunt Add-in geladen - " & Now()
End Sub

'------------------------------------------------------------------------------
' Fired when Outlook closes.
' Removes the custom context menu button.
'------------------------------------------------------------------------------
Private Sub Application_Quit()
    JobHunt.RemoveContextMenu
End Sub

'------------------------------------------------------------------------------
' Fired when a new email arrives in the Inbox.
' Auto-categorises messages that look like job-related emails.
'------------------------------------------------------------------------------
Private Sub m_Inbox_ItemAdd(ByVal Item As Object)
    If Item.Class <> olMail Then Exit Sub

    Dim oMail As Outlook.MailItem
    Set oMail = Item

    Dim subject As String
    subject = LCase(oMail.Subject)

    ' Auto-tag obvious job application responses
    If InStr(subject, "application") > 0 Or _
       InStr(subject, "vacancy") > 0 Or _
       InStr(subject, "position") > 0 Or _
       InStr(subject, "job offer") > 0 Then

        If Len(oMail.Categories) = 0 Then
            oMail.Categories = JobHunt.CATEGORY_JOB_APP
            oMail.Save
        End If
    End If
End Sub
