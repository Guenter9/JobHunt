VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JobApplicationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' Class      : JobApplicationList
' Description: Persistent collection of all JobApplication objects.
'              Stored as a CSV file (jobapplications.csv) in Settings.RootVerzeichnis.
'
' Usage:
'   Dim oList As New JobApplicationList
'   oList.Load
'   Dim oApp As JobApplication
'   Set oApp = oList.GetByTag("JA-001-AcmeEntwickler")
'   oList.Save
'==============================================================================
Option Explicit

Private Const CSV_FILE    As String = "jobapplications.csv"
Private Const CSV_HEADERS As String = "ID,Tag,Firma,Position,Ansprechpartner,Anzeige_Link," & _
                                      "Anzeige_Text,IMAPFolder,FileFoldername,Datum_Bewerbung," & _
                                      "Status,Historie,Datum_Aenderung,Notizen,Mailaccount"

Private m_Items   As Scripting.Dictionary   ' Dictionary(Tag -> JobApplication)
Private m_CsvPath As String

' ---- Initialise / Terminate --------------------------------------------------

Private Sub Class_Initialize()
    Set m_Items = New Scripting.Dictionary
    m_Items.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
    Set m_Items = Nothing
End Sub

' ---- Properties --------------------------------------------------------------

Public Property Get Count() As Long
    Count = m_Items.Count
End Property

Public Property Get CsvPath() As String
    CsvPath = m_CsvPath
End Property

' ---- Persistence -------------------------------------------------------------

'------------------------------------------------------------------------------
' Loads the CSV file from RootVerzeichnis.  Safe to call even if the file
' does not yet exist (starts with an empty collection).
'------------------------------------------------------------------------------
Public Sub Load()
    Dim sRoot As String
    sRoot = Settings.RootVerzeichnis()
    m_CsvPath = sRoot & "\" & CSV_FILE

    Set m_Items = New Scripting.Dictionary
    m_Items.CompareMode = vbTextCompare

    If Len(Dir(m_CsvPath)) = 0 Then Exit Sub       ' file not found – start empty

    ' --- Read file via ADODB.Stream (handles UTF-8 correctly)
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type    = 2    ' adTypeText
    oStream.CharSet = "utf-8"
    oStream.Open
    oStream.LoadFromFile m_CsvPath
    Dim sCsv As String
    sCsv = oStream.ReadText
    oStream.Close
    Set oStream = Nothing

    If Len(Trim(sCsv)) = 0 Then Exit Sub

    ' --- Normalise line endings and split into lines
    sCsv = Replace(sCsv, vbCrLf, vbLf)
    sCsv = Replace(sCsv, vbCr, vbLf)
    Dim aLines() As String
    aLines = Split(sCsv, vbLf)

    ' --- Line 0 is the header – data starts at line 1
    Dim i       As Long
    Dim sLine   As String
    Dim aFields() As String
    Dim oApp    As JobApplication

    For i = 1 To UBound(aLines)
        sLine = Trim(aLines(i))
        If Len(sLine) > 0 Then
            aFields = ParseCsvLine(sLine)
            If UBound(aFields) >= 14 Then
                Set oApp = New JobApplication
                oApp.FromCsvFields aFields
                If Len(oApp.Tag) > 0 Then
                    Set m_Items(oApp.Tag) = oApp
                End If
            End If
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' Serialises the collection to CSV and writes the file.
'------------------------------------------------------------------------------
Public Sub Save()
    Dim sRoot As String
    sRoot = Settings.RootVerzeichnis()

    ' Ensure root directory exists
    If Len(Dir(sRoot, vbDirectory)) = 0 Then
        MkDir sRoot
    End If

    m_CsvPath = sRoot & "\" & CSV_FILE

    ' --- Build CSV content
    Dim sOut    As String
    Dim oApp    As JobApplication
    Dim aFields() As String
    Dim i       As Long
    Dim sRow    As String

    sOut = CSV_HEADERS & vbNewLine

    Dim vKey As Variant
    For Each vKey In m_Items.Keys
        Set oApp = m_Items(vKey)
        aFields = oApp.ToCsvFields()
        sRow = ""
        For i = 0 To UBound(aFields)
            If i > 0 Then sRow = sRow & ","
            sRow = sRow & QuoteCsvField(aFields(i))
        Next i
        sOut = sOut & sRow & vbNewLine
    Next vKey

    ' --- Write via ADODB.Stream (UTF-8 without BOM)
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type    = 2
    oStream.CharSet = "utf-8"
    oStream.Open
    oStream.WriteText sOut
    oStream.Position = 0
    ' Remove UTF-8 BOM (first 3 bytes)
    oStream.Type = 1   ' adTypeBinary
    oStream.Position = 3
    Dim oBomFree As Object
    Set oBomFree = CreateObject("ADODB.Stream")
    oBomFree.Type = 1
    oBomFree.Open
    oStream.CopyTo oBomFree
    oBomFree.SaveToFile m_CsvPath, 2   ' adSaveCreateOverWrite
    oBomFree.Close
    oStream.Close
    Set oStream = Nothing
    Set oBomFree = Nothing
End Sub

' ---- Collection management ---------------------------------------------------

'------------------------------------------------------------------------------
' Adds a JobApplication to the dictionary.
' If an item with the same Tag already exists it is replaced.
'------------------------------------------------------------------------------
Public Sub Add(oApp As JobApplication)
    Set m_Items(oApp.Tag) = oApp
End Sub

'------------------------------------------------------------------------------
' Removes the JobApplication with the given Tag.
'------------------------------------------------------------------------------
Public Sub Remove(sTag As String)
    If m_Items.Exists(sTag) Then m_Items.Remove sTag
End Sub

'------------------------------------------------------------------------------
' Returns the JobApplication whose Tag matches, or Nothing if not found.
'------------------------------------------------------------------------------
Public Function GetByTag(sTag As String) As JobApplication
    If m_Items.Exists(sTag) Then Set GetByTag = m_Items(sTag)
End Function

'------------------------------------------------------------------------------
' Returns the JobApplication whose ID matches, or Nothing if not found.
'------------------------------------------------------------------------------
Public Function GetByID(sID As String) As JobApplication
    Dim vKey As Variant
    Dim oApp As JobApplication
    For Each vKey In m_Items.Keys
        Set oApp = m_Items(vKey)
        If oApp.ID = sID Then
            Set GetByID = oApp
            Exit Function
        End If
    Next vKey
End Function

'------------------------------------------------------------------------------
' Returns the underlying Dictionary (Tag -> JobApplication) for iteration.
'------------------------------------------------------------------------------
Public Function GetAll() As Scripting.Dictionary
    Set GetAll = m_Items
End Function

' ---- CSV helpers -------------------------------------------------------------

'------------------------------------------------------------------------------
' Quotes a single CSV field value.
' Embedded newlines are encoded as \n to keep every record on one line.
' Fields containing commas or double-quotes are wrapped in double-quotes.
'------------------------------------------------------------------------------
Private Function QuoteCsvField(s As String) As String
    ' Encode embedded newlines as the two-character sequence \n
    Dim sOut As String
    sOut = Replace(Replace(Replace(s, vbCrLf, "\n"), vbCr, "\n"), vbLf, "\n")
    ' Wrap in double-quotes when the field contains a comma or a double-quote
    If InStr(sOut, ",") > 0 Or InStr(sOut, """") > 0 Then
        QuoteCsvField = """" & Replace(sOut, """", """""") & """"
    Else
        QuoteCsvField = sOut
    End If
End Function

'------------------------------------------------------------------------------
' Parses one CSV line into a zero-based String array.
' Handles RFC 4180 quoted fields; decodes \n back to vbNewLine.
'------------------------------------------------------------------------------
Private Function ParseCsvLine(sLine As String) As String()
    Dim aOut(0 To 20) As String
    Dim nIdx    As Long
    Dim i       As Long
    Dim ch      As String
    Dim sField  As String
    Dim bQuoted As Boolean

    nIdx    = 0
    i       = 1
    sField  = ""
    bQuoted = False

    Do While i <= Len(sLine)
        ch = Mid(sLine, i, 1)
        If bQuoted Then
            If ch = """" Then
                If Mid(sLine, i + 1, 1) = """"  Then  ' escaped double-quote
                    sField = sField & """"
                    i = i + 1
                Else
                    bQuoted = False                    ' closing quote
                End If
            Else
                sField = sField & ch
            End If
        Else
            Select Case ch
                Case """"
                    bQuoted = True
                Case ","
                    aOut(nIdx) = Replace(sField, "\n", vbNewLine)
                    nIdx   = nIdx + 1
                    sField = ""
                Case Else
                    sField = sField & ch
            End Select
        End If
        i = i + 1
    Loop

    ' Store last field
    aOut(nIdx) = Replace(sField, "\n", vbNewLine)
    nIdx = nIdx + 1

    ReDim Preserve aOut(0 To nIdx - 1)
    ParseCsvLine = aOut
End Function

'------------------------------------------------------------------------------
' Returns the next available zero-padded 3-digit ID string (e.g. "042").
'------------------------------------------------------------------------------
Public Function NextID() As String
    Dim lMax As Long
    lMax = 0
    Dim vKey As Variant
    Dim oApp As JobApplication
    For Each vKey In m_Items.Keys
        Set oApp = m_Items(vKey)
        Dim lCur As Long
        On Error Resume Next
        lCur = CLng(oApp.ID)
        On Error GoTo 0
        If lCur > lMax Then lMax = lCur
    Next vKey
    NextID = Format(lMax + 1, "000")
End Function
