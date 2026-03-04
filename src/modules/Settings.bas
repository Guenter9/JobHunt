Attribute VB_Name = "Settings"
'==============================================================================
' Module     : Settings
' Description: Reads and writes application settings from/to an INI-style file.
'
' Config file location:  %APPDATA%\JobHunt\settings.ini
'
' Keys (case-insensitive):
'   RootVerzeichnis  - file-system root folder for JobApplication sub-folders
'   ImapRoot         - display name of the IMAP parent folder (e.g. "Bewerbungen")
'   VorlageMail      - full path to an .oft mail template file
'
' Usage (read):
'   Dim sRoot As String
'   sRoot = Settings.RootVerzeichnis()
'
' Usage (write):
'   Settings.SetRootVerzeichnis "C:\Users\me\Bewerbungen"
'==============================================================================
Option Explicit

Private Const APP_FOLDER    As String = "JobHunt"
Private Const CONFIG_FILE   As String = "settings.ini"

Private Const KEY_ROOT      As String = "RootVerzeichnis"
Private Const KEY_IMAP      As String = "ImapRoot"
Private Const KEY_TEMPLATE  As String = "VorlageMail"

' ---- Convenience accessors ---------------------------------------------------

Public Function RootVerzeichnis() As String
    RootVerzeichnis = GetSetting(KEY_ROOT)
End Function

Public Sub SetRootVerzeichnis(sValue As String)
    PutSetting KEY_ROOT, sValue
End Sub

Public Function ImapRoot() As String
    ImapRoot = GetSetting(KEY_IMAP)
End Function

Public Sub SetImapRoot(sValue As String)
    PutSetting KEY_IMAP, sValue
End Sub

Public Function VorlageMail() As String
    VorlageMail = GetSetting(KEY_TEMPLATE)
End Function

Public Sub SetVorlageMail(sValue As String)
    PutSetting KEY_TEMPLATE, sValue
End Sub

' ---- Core read/write ---------------------------------------------------------

'------------------------------------------------------------------------------
' Returns the value for the given key, or an empty string if not found.
'------------------------------------------------------------------------------
Public Function GetSetting(sKey As String) As String
    Dim sPath As String
    sPath = ConfigFilePath()
    If Not FileExists(sPath) Then
        GetSetting = ""
        Exit Function
    End If

    Dim iFile As Integer
    iFile = FreeFile
    Open sPath For Input As #iFile
    Dim sLine As String
    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        sLine = Trim(sLine)
        If Left(sLine, 1) = ";" Or Len(sLine) = 0 Then GoTo NextLine  ' comment / blank
        Dim iEq As Integer
        iEq = InStr(sLine, "=")
        If iEq > 0 Then
            Dim sK As String
            sK = Trim(Left(sLine, iEq - 1))
            If LCase(sK) = LCase(sKey) Then
                Close #iFile
                GetSetting = Trim(Mid(sLine, iEq + 1))
                Exit Function
            End If
        End If
NextLine:
    Loop
    Close #iFile
    GetSetting = ""
End Function

'------------------------------------------------------------------------------
' Writes (or updates) key=value in the config file.
'------------------------------------------------------------------------------
Public Sub PutSetting(sKey As String, sValue As String)
    EnsureConfigDir

    Dim sPath As String
    sPath = ConfigFilePath()

    ' Read all lines into an array
    Dim aLines() As String
    Dim nLines As Long
    nLines = 0

    If FileExists(sPath) Then
        Dim iFile As Integer
        iFile = FreeFile
        Open sPath For Input As #iFile
        Dim sLine As String
        Do While Not EOF(iFile)
            Line Input #iFile, sLine
            ReDim Preserve aLines(nLines)
            aLines(nLines) = sLine
            nLines = nLines + 1
        Loop
        Close #iFile
    End If

    ' Find and replace existing key, or flag as not found
    Dim bFound As Boolean
    bFound = False
    Dim i As Long
    For i = 0 To nLines - 1
        Dim iEq As Integer
        iEq = InStr(aLines(i), "=")
        If iEq > 0 Then
            If LCase(Trim(Left(aLines(i), iEq - 1))) = LCase(sKey) Then
                aLines(i) = sKey & "=" & sValue
                bFound = True
                Exit For
            End If
        End If
    Next i

    If Not bFound Then
        ReDim Preserve aLines(nLines)
        aLines(nLines) = sKey & "=" & sValue
        nLines = nLines + 1
    End If

    ' Write all lines back
    Dim oFile As Integer
    oFile = FreeFile
    Open sPath For Output As #oFile
    For i = 0 To nLines - 1
        Print #oFile, aLines(i)
    Next i
    Close #oFile
End Sub

'------------------------------------------------------------------------------
' Called at startup: creates the config directory and a default config file
' if they do not already exist.
'------------------------------------------------------------------------------
Public Sub EnsureSettingsExist()
    EnsureConfigDir
    Dim sPath As String
    sPath = ConfigFilePath()
    If Not FileExists(sPath) Then
        Dim iFile As Integer
        iFile = FreeFile
        Open sPath For Output As #iFile
        Print #iFile, "; JobHunt settings"
        Print #iFile, "; Created: " & Now()
        Print #iFile, ""
        Print #iFile, "; Root folder where sub-folders for each JobApplication are created"
        Print #iFile, KEY_ROOT & "=" & Environ("USERPROFILE") & "\Documents\Bewerbungen"
        Print #iFile, ""
        Print #iFile, "; Display name of the IMAP parent folder for JobApplication sub-folders"
        Print #iFile, KEY_IMAP & "=Bewerbungen"
        Print #iFile, ""
        Print #iFile, "; Full path to an Outlook template file (.oft) – leave blank to skip"
        Print #iFile, KEY_TEMPLATE & "="
        Close #iFile
        MsgBox "JobHunt: Standardeinstellungen wurden angelegt." & vbNewLine & _
               "Bitte prüfe die Einstellungsdatei:" & vbNewLine & sPath, _
               vbInformation, "JobHunt – Einstellungen"
    End If
End Sub

'------------------------------------------------------------------------------
' Opens the settings.ini file in the default text editor.
'------------------------------------------------------------------------------
Public Sub OpenSettingsFile()
    EnsureSettingsExist
    Shell "notepad.exe """ & ConfigFilePath() & """", vbNormalFocus
End Sub

' ---- Path helpers ------------------------------------------------------------

Public Function ConfigFilePath() As String
    ConfigFilePath = ConfigDirPath() & "\" & CONFIG_FILE
End Function

Public Function ConfigDirPath() As String
    ConfigDirPath = Environ("APPDATA") & "\" & APP_FOLDER
End Function

Private Sub EnsureConfigDir()
    Dim sDir As String
    sDir = ConfigDirPath()
    If Not FolderExists(sDir) Then
        MkDir sDir
    End If
End Sub

Private Function FileExists(sPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(sPath)) > 0)
    On Error GoTo 0
End Function

Private Function FolderExists(sPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Len(Dir(sPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function
