VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : JobApplication
' Description: Represents a single job application.
'              Manages file-system folder, IMAP folder, Outlook categories,
'              and a history log.
'
' Typical flow:
'   Dim oApp As New JobApplication
'   oApp.Init "Acme GmbH", "Softwareentwickler"
'   oApp.Create                          ' creates folders, registers in list
'   oApp.AddMail ActiveInspector.CurrentItem, "gesendet", "Bewerbung abgeschickt"
'==============================================================================
Option Explicit

' ---- Private fields ----------------------------------------------------------
Private m_ID               As String
Private m_Tag              As String
Private m_Firma            As String
Private m_Position         As String
Private m_Ansprechpartner  As String
Private m_Anzeige_Link     As String
Private m_Anzeige_Text     As String
Private m_IMAPFolder       As String
Private m_FileFoldername   As String
Private m_Datum_Bewerbung  As Date
Private m_Status           As String
Private m_Historie         As String
Private m_Datum_Aenderung  As Date
Private m_Notizen          As String
Private m_Mailaccount      As String

' ---- Properties --------------------------------------------------------------

Public Property Get ID() As String:               ID = m_ID:               End Property
Public Property Let ID(v As String):              m_ID = v:                End Property

Public Property Get Tag() As String:              Tag = m_Tag:             End Property
Public Property Let Tag(v As String):             m_Tag = v:               End Property

Public Property Get Firma() As String:            Firma = m_Firma:         End Property
Public Property Let Firma(v As String):           m_Firma = v:             End Property

Public Property Get Position() As String:         Position = m_Position:   End Property
Public Property Let Position(v As String):        m_Position = v:          End Property

Public Property Get Ansprechpartner() As String:  Ansprechpartner = m_Ansprechpartner: End Property
Public Property Let Ansprechpartner(v As String): m_Ansprechpartner = v:   End Property

Public Property Get Anzeige_Link() As String:     Anzeige_Link = m_Anzeige_Link: End Property
Public Property Let Anzeige_Link(v As String):    m_Anzeige_Link = v:      End Property

Public Property Get Anzeige_Text() As String:     Anzeige_Text = m_Anzeige_Text: End Property
Public Property Let Anzeige_Text(v As String):    m_Anzeige_Text = v:      End Property

Public Property Get IMAPFolder() As String:       IMAPFolder = m_IMAPFolder: End Property
Public Property Let IMAPFolder(v As String):      m_IMAPFolder = v:        End Property

Public Property Get FileFoldername() As String:   FileFoldername = m_FileFoldername: End Property
Public Property Let FileFoldername(v As String):  m_FileFoldername = v:    End Property

Public Property Get Datum_Bewerbung() As Date:    Datum_Bewerbung = m_Datum_Bewerbung: End Property
Public Property Let Datum_Bewerbung(v As Date):   m_Datum_Bewerbung = v:   End Property

Public Property Get Status() As String:           Status = m_Status:       End Property
Public Property Let Status(v As String):          m_Status = v:            End Property

Public Property Get Historie() As String:         Historie = m_Historie:   End Property
Public Property Let Historie(v As String):        m_Historie = v:          End Property

Public Property Get Datum_Aenderung() As Date:    Datum_Aenderung = m_Datum_Aenderung: End Property
Public Property Let Datum_Aenderung(v As Date):   m_Datum_Aenderung = v:   End Property

Public Property Get Notizen() As String:          Notizen = m_Notizen:     End Property
Public Property Let Notizen(v As String):         m_Notizen = v:           End Property

Public Property Get Mailaccount() As String:      Mailaccount = m_Mailaccount: End Property
Public Property Let Mailaccount(v As String):     m_Mailaccount = v:       End Property

' ---- Main Methods ------------------------------------------------------------

'------------------------------------------------------------------------------
' Init: Initialises a new JobApplication with mandatory Firma and Position.
'       Generates folder names, Tag, and auto-increments the ID.
'
'  Parameters:
'    sFirma    - Company name  (required)
'    sPosition - Job title     (required)
'------------------------------------------------------------------------------
Public Sub Init(sFirma As String, sPosition As String)
    If Len(Trim(sFirma)) = 0 Or Len(Trim(sPosition)) = 0 Then
        Err.Raise vbObjectError + 1001, "JobApplication.Init", _
                  "Firma und Position sind Pflichtfelder."
    End If

    m_Firma    = Trim(sFirma)
    m_Position = Trim(sPosition)
    m_Status   = "geplant"

    ' Auto-increment ID from the global list
    Dim oList As New JobApplicationList
    oList.Load
    m_ID = oList.NextID()

    ' Generate folder / IMAP names
    Dim sSafe As String
    sSafe = SanitizeName(m_Firma & "_" & m_Position)

    m_FileFoldername = m_ID & "_" & sSafe
    m_IMAPFolder     = m_ID & "_" & sSafe

    ' Build Tag: "JA-NNN-FirmaPos" (max 50 chars total)
    m_Tag = "JA-" & m_ID & "-" & SanitizeName(m_Firma & m_Position)
    If Len(m_Tag) > 50 Then m_Tag = Left(m_Tag, 50)
End Sub

'------------------------------------------------------------------------------
' Create: Persists the application.
'   - Creates the file-system folder under Settings.RootVerzeichnis
'   - Creates the IMAP sub-folder under Settings.ImapRoot
'   - Registers the object in JobApplicationList and saves it
'   - Writes the first history entry "erzeugt"
'------------------------------------------------------------------------------
Public Sub Create()
    If Len(m_Firma) = 0 Or Len(m_Position) = 0 Then
        Err.Raise vbObjectError + 1002, "JobApplication.Create", _
                  "Init() muss vor Create() aufgerufen werden."
    End If

    ' --- 1. File-system folder
    Dim sFileDir As String
    sFileDir = Settings.RootVerzeichnis() & "\" & m_FileFoldername
    If Len(Dir(sFileDir, vbDirectory)) = 0 Then
        MkDir sFileDir
    End If

    ' --- 2. IMAP folder
    CreateIMAPFolder

    ' --- 3. History entry
    AppendHistorie "erzeugt"

    ' --- 4. Register and persist
    Dim oList As New JobApplicationList
    oList.Load
    oList.Add Me
    oList.Save
End Sub

'------------------------------------------------------------------------------
' AddMail: Associates a MailItem with this application.
'   - Tags the mail with the application's Outlook category
'   - Saves a local .msg copy to the file folder
'   - Moves the mail to the IMAP folder
'   - Updates status & history, persists the list
'
'  Parameters:
'    aMailItem  - The Outlook MailItem to process
'    aStatus    - New status (geplant|gesendet|aktiv|archiviert)
'    aVorgang   - Short description for the history log
'------------------------------------------------------------------------------
Public Sub AddMail(aMailItem As Outlook.MailItem, _
                   aStatus   As String, _
                   aVorgang  As String)

    If aMailItem Is Nothing Then
        Err.Raise vbObjectError + 1003, "JobApplication.AddMail", _
                  "aMailItem darf nicht Nothing sein."
    End If

    ' --- 1. Apply Outlook category tag
    Dim sCats As String
    sCats = aMailItem.Categories
    If InStr(sCats, m_Tag) = 0 Then
        If Len(sCats) > 0 Then
            aMailItem.Categories = sCats & ", " & m_Tag
        Else
            aMailItem.Categories = m_Tag
        End If
    End If

    ' --- 2. Detect mail account from store
    On Error Resume Next
    Dim sAccount As String
    sAccount = aMailItem.Parent.Store.DisplayName
    On Error GoTo 0
    If Len(sAccount) > 0 Then m_Mailaccount = sAccount

    ' --- 3. Save local .msg copy
    Dim sFileDir As String
    sFileDir = Settings.RootVerzeichnis() & "\" & m_FileFoldername
    If Len(Dir(sFileDir, vbDirectory)) = 0 Then MkDir sFileDir

    Dim sSafeSubject As String
    sSafeSubject = SanitizeName(aMailItem.Subject)
    Dim sFilePath As String
    sFilePath = sFileDir & "\" & _
                Format(Now, "YYYYMMDD_HHMMSS") & "_" & sSafeSubject & ".msg"
    aMailItem.SaveAs sFilePath, olMSG

    ' --- 4. Move to IMAP folder
    Dim oImapFolder As Outlook.MAPIFolder
    Set oImapFolder = GetIMAPFolder()
    If Not oImapFolder Is Nothing Then
        aMailItem.Move oImapFolder
    Else
        aMailItem.Save   ' at least save the category tag
    End If

    ' --- 5. Update metadata
    m_Status          = aStatus
    m_Datum_Aenderung = Now
    AppendHistorie aVorgang

    ' --- 6. Persist
    Dim oList As New JobApplicationList
    oList.Load
    oList.Add Me
    oList.Save
End Sub

'------------------------------------------------------------------------------
' UpdateAndSave: Persists changed metadata (Status, Notizen, Vorgang) without
'               processing a mail.
'
'  Parameters:
'    aStatus  - New status (may be empty to keep current)
'    aVorgang - History entry (may be empty to skip)
'------------------------------------------------------------------------------
Public Sub UpdateAndSave(Optional aStatus As String = "", _
                         Optional aVorgang As String = "")
    If Len(aStatus) > 0 Then m_Status = aStatus
    If Len(aVorgang) > 0 Then AppendHistorie aVorgang
    m_Datum_Aenderung = Now

    Dim oList As New JobApplicationList
    oList.Load
    oList.Add Me
    oList.Save
End Sub

' ---- Serialization -----------------------------------------------------------

'------------------------------------------------------------------------------
' ToDict: Returns a Scripting.Dictionary representation for JSON serialization.
'------------------------------------------------------------------------------
Public Function ToDict() As Scripting.Dictionary
    Dim d As New Scripting.Dictionary
    d("ID")               = m_ID
    d("Tag")              = m_Tag
    d("Firma")            = m_Firma
    d("Position")         = m_Position
    d("Ansprechpartner")  = m_Ansprechpartner
    d("Anzeige_Link")     = m_Anzeige_Link
    d("Anzeige_Text")     = m_Anzeige_Text
    d("IMAPFolder")       = m_IMAPFolder
    d("FileFoldername")   = m_FileFoldername
    d("Datum_Bewerbung")  = IIf(m_Datum_Bewerbung = 0, "", _
                               Format(m_Datum_Bewerbung, "YYYY-MM-DD"))
    d("Status")           = m_Status
    d("Historie")         = m_Historie
    d("Datum_Aenderung")  = IIf(m_Datum_Aenderung = 0, "", _
                               Format(m_Datum_Aenderung, "YYYY-MM-DD HH:MM:SS"))
    d("Notizen")          = m_Notizen
    d("Mailaccount")      = m_Mailaccount
    Set ToDict = d
End Function

'------------------------------------------------------------------------------
' FromDict: Populates this instance from a Scripting.Dictionary (JSON object).
'------------------------------------------------------------------------------
Public Sub FromDict(d As Scripting.Dictionary)
    m_ID              = DictStr(d, "ID")
    m_Tag             = DictStr(d, "Tag")
    m_Firma           = DictStr(d, "Firma")
    m_Position        = DictStr(d, "Position")
    m_Ansprechpartner = DictStr(d, "Ansprechpartner")
    m_Anzeige_Link    = DictStr(d, "Anzeige_Link")
    m_Anzeige_Text    = DictStr(d, "Anzeige_Text")
    m_IMAPFolder      = DictStr(d, "IMAPFolder")
    m_FileFoldername  = DictStr(d, "FileFoldername")
    m_Status          = DictStr(d, "Status")
    m_Historie        = DictStr(d, "Historie")
    m_Notizen         = DictStr(d, "Notizen")
    m_Mailaccount     = DictStr(d, "Mailaccount")

    On Error Resume Next
    Dim sDatBew As String
    sDatBew = DictStr(d, "Datum_Bewerbung")
    If Len(sDatBew) > 0 Then m_Datum_Bewerbung = CDate(sDatBew)
    Dim sDatAend As String
    sDatAend = DictStr(d, "Datum_Aenderung")
    If Len(sDatAend) > 0 Then m_Datum_Aenderung = CDate(sDatAend)
    On Error GoTo 0
End Sub

' ---- Private Helpers ---------------------------------------------------------

'------------------------------------------------------------------------------
' Prepends a timestamped entry to the Historie memo.
'------------------------------------------------------------------------------
Private Sub AppendHistorie(sVorgang As String)
    Dim sEntry As String
    sEntry = Format(Now, "YYYY-MM-DD HH:MM:SS") & " - " & sVorgang
    If Len(m_Historie) > 0 Then
        m_Historie = sEntry & vbNewLine & m_Historie
    Else
        m_Historie = sEntry
    End If
End Sub

'------------------------------------------------------------------------------
' Creates the IMAP sub-folder under the configured ImapRoot folder, if it
' does not already exist.  Iterates all registered mail stores to find one
' containing ImapRoot.
'------------------------------------------------------------------------------
Private Sub CreateIMAPFolder()
    Dim oFolder As Outlook.MAPIFolder
    Set oFolder = GetIMAPRootFolder()
    If oFolder Is Nothing Then Exit Sub   ' ImapRoot not found - skip silently

    On Error Resume Next
    Dim oSub As Outlook.MAPIFolder
    Set oSub = oFolder.Folders(m_IMAPFolder)
    On Error GoTo 0

    If oSub Is Nothing Then
        oFolder.Folders.Add m_IMAPFolder
    End If
End Sub

'------------------------------------------------------------------------------
' Returns the IMAP sub-folder for this application, or Nothing.
'------------------------------------------------------------------------------
Private Function GetIMAPFolder() As Outlook.MAPIFolder
    Dim oRoot As Outlook.MAPIFolder
    Set oRoot = GetIMAPRootFolder()
    If oRoot Is Nothing Then Exit Function

    On Error Resume Next
    Set GetIMAPFolder = oRoot.Folders(m_IMAPFolder)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Finds the ImapRoot folder across all mail stores.
'------------------------------------------------------------------------------
Private Function GetIMAPRootFolder() As Outlook.MAPIFolder
    Dim sImapRoot As String
    sImapRoot = Settings.ImapRoot()
    If Len(sImapRoot) = 0 Then Exit Function

    Dim oNS    As Outlook.NameSpace
    Set oNS = Application.GetNamespace("MAPI")

    Dim oStore  As Outlook.Store
    Dim oFolder As Outlook.MAPIFolder
    For Each oStore In oNS.Stores
        On Error Resume Next
        Set oFolder = oStore.GetRootFolder().Folders(sImapRoot)
        On Error GoTo 0
        If Not oFolder Is Nothing Then
            Set GetIMAPRootFolder = oFolder
            Exit Function
        End If
    Next oStore
End Function

'------------------------------------------------------------------------------
' Converts a raw string to a filesystem/IMAP-safe name:
'   - Strips characters that are invalid in folder names
'   - Replaces spaces with "_"
'   - Truncates to 40 characters
'------------------------------------------------------------------------------
Private Function SanitizeName(s As String) As String
    Dim sOut As String
    sOut = ""
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "_", "-"
                sOut = sOut & ch
            Case " ", Chr(9)
                sOut = sOut & "_"
            Case Chr(228)   ' ae
                sOut = sOut & "ae"
            Case Chr(246)   ' oe
                sOut = sOut & "oe"
            Case Chr(252)   ' ue
                sOut = sOut & "ue"
            Case Chr(196)   ' Ae
                sOut = sOut & "Ae"
            Case Chr(214)   ' Oe
                sOut = sOut & "Oe"
            Case Chr(220)   ' Ue
                sOut = sOut & "Ue"
            Case Chr(223)   ' ss
                sOut = sOut & "ss"
            ' All other chars are silently dropped
        End Select
    Next i
    ' Collapse multiple underscores
    Do While InStr(sOut, "__") > 0
        sOut = Replace(sOut, "__", "_")
    Loop
    If Len(sOut) > 40 Then sOut = Left(sOut, 40)
    SanitizeName = sOut
End Function

'------------------------------------------------------------------------------
' Safe Dictionary string reader - returns "" if key is missing.
'------------------------------------------------------------------------------
Private Function DictStr(d As Scripting.Dictionary, sKey As String) As String
    If d.Exists(sKey) Then
        On Error Resume Next
        DictStr = CStr(d(sKey))
        On Error GoTo 0
    End If
End Function
