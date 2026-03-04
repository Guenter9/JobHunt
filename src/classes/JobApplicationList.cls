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
'              Stored as a JSON file (jobapplications.json) in Settings.RootVerzeichnis.
'
' Usage:
'   Dim oList As New JobApplicationList
'   oList.Load
'   Dim oApp As JobApplication
'   Set oApp = oList.GetByTag("JA-001-AcmeEntwickler")
'   oList.Save
'==============================================================================
Option Explicit

Private Const JSON_FILE As String = "jobapplications.json"

Private m_Items    As Collection   ' Collection of JobApplication objects
Private m_JsonPath As String

' ---- Initialise / Terminate --------------------------------------------------

Private Sub Class_Initialize()
    Set m_Items = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Items = Nothing
End Sub

' ---- Properties --------------------------------------------------------------

Public Property Get Count() As Long
    Count = m_Items.Count
End Property

Public Property Get JsonPath() As String
    JsonPath = m_JsonPath
End Property

' ---- Persistence -------------------------------------------------------------

'------------------------------------------------------------------------------
' Loads the JSON file from RootVerzeichnis.  Safe to call even if the file
' does not yet exist (starts with an empty collection).
'------------------------------------------------------------------------------
Public Sub Load()
    Dim sRoot As String
    sRoot = Settings.RootVerzeichnis()
    m_JsonPath = sRoot & "\" & JSON_FILE

    Set m_Items = New Collection

    If Len(Dir(m_JsonPath)) = 0 Then Exit Sub       ' file not found – start empty

    ' --- Read file via ADODB.Stream (handles encoding correctly)
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type    = 2    ' adTypeText
    oStream.CharSet = "utf-8"
    oStream.Open
    oStream.LoadFromFile m_JsonPath
    Dim sJson As String
    sJson = oStream.ReadText
    oStream.Close
    Set oStream = Nothing

    If Len(Trim(sJson)) = 0 Then Exit Sub

    ' --- Parse JSON
    Dim oRoot As Object
    Set oRoot = JsonConverter.ParseJson(sJson)

    ' Root should be a Collection (JSON array) or it could be a Dict with key "value"
    Dim oArray As Collection
    If TypeName(oRoot) = "Collection" Then
        Set oArray = oRoot
    ElseIf TypeName(oRoot) = "Dictionary" Then
        ' Fallback: wrapped by ParseJson scalar handler
        Exit Sub
    End If

    ' --- Rebuild JobApplication objects
    Dim vItem As Variant
    For Each vItem In oArray
        If TypeName(vItem) = "Dictionary" Then
            Dim oApp As New JobApplication
            oApp.FromDict vItem
            m_Items.Add oApp, oApp.Tag
        End If
    Next vItem
End Sub

'------------------------------------------------------------------------------
' Serializes the collection to JSON and writes the file.
'------------------------------------------------------------------------------
Public Sub Save()
    Dim sRoot As String
    sRoot = Settings.RootVerzeichnis()

    ' Ensure root directory exists
    If Len(Dir(sRoot, vbDirectory)) = 0 Then
        MkDir sRoot
    End If

    m_JsonPath = sRoot & "\" & JSON_FILE

    ' Build a Collection of Dictionaries
    Dim oArray As New Collection
    Dim oApp   As JobApplication
    For Each oApp In m_Items
        oArray.Add oApp.ToDict()
    Next oApp

    Dim sJson As String
    sJson = JsonConverter.ConvertToJson(oArray, 2)

    ' Write via ADODB.Stream (UTF-8 without BOM)
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type    = 2
    oStream.CharSet = "utf-8"
    oStream.Open
    oStream.WriteText sJson
    oStream.Position = 0
    ' Remove UTF-8 BOM (first 3 bytes)
    oStream.Type = 1   ' adTypeBinary
    oStream.Position = 3
    Dim oBomFree As Object
    Set oBomFree = CreateObject("ADODB.Stream")
    oBomFree.Type = 1
    oBomFree.Open
    oStream.CopyTo oBomFree
    oBomFree.SaveToFile m_JsonPath, 2   ' adSaveCreateOverWrite
    oBomFree.Close
    oStream.Close
    Set oStream = Nothing
    Set oBomFree = Nothing
End Sub

' ---- Collection management ---------------------------------------------------

'------------------------------------------------------------------------------
' Adds a JobApplication to the collection.
' If an item with the same Tag already exists it is replaced.
'------------------------------------------------------------------------------
Public Sub Add(oApp As JobApplication)
    ' Remove existing entry with same tag to avoid duplicate key error
    On Error Resume Next
    m_Items.Remove oApp.Tag
    On Error GoTo 0
    m_Items.Add oApp, oApp.Tag
End Sub

'------------------------------------------------------------------------------
' Removes the JobApplication with the given Tag.
'------------------------------------------------------------------------------
Public Sub Remove(sTag As String)
    On Error Resume Next
    m_Items.Remove sTag
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Returns the JobApplication whose Tag matches, or Nothing if not found.
'------------------------------------------------------------------------------
Public Function GetByTag(sTag As String) As JobApplication
    On Error Resume Next
    Set GetByTag = m_Items(sTag)
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Returns the JobApplication whose ID matches, or Nothing if not found.
'------------------------------------------------------------------------------
Public Function GetByID(sID As String) As JobApplication
    Dim oApp As JobApplication
    For Each oApp In m_Items
        If oApp.ID = sID Then
            Set GetByID = oApp
            Exit Function
        End If
    Next oApp
End Function

'------------------------------------------------------------------------------
' Returns the full Collection for iteration.
'------------------------------------------------------------------------------
Public Function GetAll() As Collection
    Set GetAll = m_Items
End Function

'------------------------------------------------------------------------------
' Returns the next available zero-padded 3-digit ID string (e.g. "042").
'------------------------------------------------------------------------------
Public Function NextID() As String
    Dim lMax As Long
    lMax = 0
    Dim oApp As JobApplication
    For Each oApp In m_Items
        Dim lCur As Long
        On Error Resume Next
        lCur = CLng(oApp.ID)
        On Error GoTo 0
        If lCur > lMax Then lMax = lCur
    Next oApp
    NextID = Format(lMax + 1, "000")
End Function
