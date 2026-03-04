VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJobApplication 
   Caption         =   "Job Application verwalten"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJobApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' Form       : frmJobApplication
' Description: Job-Application-Verwaltungsformular.
'
'  Alle Controls werden zur Laufzeit per Me.Controls.Add erstellt –
'  kein .frx-File und kein Application.VBE-Zugriff notwenidg.
'  Events werden ueber clsBtnSink / clsOptSink / clsCboSink weitergeleitet.
'
'  Import: VBA-Editor (Alt+F11) > Datei > Importieren > frmJobApplication.frm
'  Voraussetzung: clsBtnSink, clsOptSink, clsCboSink bereits importiert.
'==============================================================================
Option Explicit

' ---- Data -------------------------------------------------------------------
Private m_Mail As Outlook.MailItem
Private m_App  As JobApplication
Private m_List As JobApplicationList

' ---- Event-sink collection (keeps sinks alive) ------------------------------
Private m_Sinks As Collection

' =============================================================================
'  Initialization
' =============================================================================

Private Sub UserForm_Initialize()
    BuildControls
    InitControlProps
    WireEvents
    LoadData
End Sub

' ---- 1. Create every control with geometry only ----------------------------

Private Sub BuildControls()
    Dim L1 As Single: L1 = 6       ' label left edge
    Dim L2 As Single: L2 = 114     ' input left edge
    Dim W1 As Single: W1 = 102     ' label width
    Dim W2 As Single: W2 = 268     ' input/combo width
    Dim H  As Single: H  = 16      ' single-line height
    Dim HM As Single: HM = 54      ' multiline height
    Dim T  As Single: T  = 4       ' running vertical position

    ' Mail info labels
    AddCtl "Forms.Label.1",         "lblMailHdr",         L1, T, W1+W2+6, H  : T = T + H + 2
    AddCtl "Forms.Label.1",         "lblMailFrom",        L1, T, W1+W2+6, H  : T = T + H + 2
    AddCtl "Forms.Label.1",         "lblMailSubject",     L1, T, W1+W2+6, H  : T = T + H + 4
    ' Separator 1
    AddCtl "Forms.Label.1",         "lblSep1",            L1, T, W1+W2+6, 2  : T = T + 5
    ' Neu / Bestehend row
    AddCtl "Forms.OptionButton.1",  "optNeu",             L1,  T, 102,    H
    AddCtl "Forms.OptionButton.1",  "optBestehend",       114, T, 100,    H
    AddCtl "Forms.ComboBox.1",      "cboJobApplications", 220, T, 162,    H  : T = T + H + 4
    ' Separator 2
    AddCtl "Forms.Label.1",         "lblSep2",            L1, T, W1+W2+6, 2  : T = T + 5
    ' Input fields
    AddCtl "Forms.Label.1",         "lblFirma",           L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtFirma",           L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblPosition",        L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtPosition",        L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblAnsprechpartner", L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtAnsprechpartner", L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblAnzeigeLink",     L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtAnzeigeLink",     L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblAnzeigeText",     L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtAnzeigeText",     L2, T, W2,      HM : T = T + HM + 3
    AddCtl "Forms.Label.1",         "lblStatus",          L1, T, W1,      H
    AddCtl "Forms.ComboBox.1",      "cboStatus",          L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblVorgang",         L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtVorgang",         L2, T, W2,      H  : T = T + H + 3
    AddCtl "Forms.Label.1",         "lblNotizen",         L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtNotizen",         L2, T, W2,      HM : T = T + HM + 3
    AddCtl "Forms.Label.1",         "lblHistorie",        L1, T, W1,      H
    AddCtl "Forms.TextBox.1",       "txtHistorie",        L2, T, W2,      HM : T = T + HM + 6
    ' Buttons
    AddCtl "Forms.CommandButton.1", "btnZuordnen",        L1,  T, 110,    22
    AddCtl "Forms.CommandButton.1", "btnSpeichern",       122, T, 110,    22
    AddCtl "Forms.CommandButton.1", "btnClose",           238, T,  76,    22
    AddCtl "Forms.CommandButton.1", "btnSettings",        320, T,  56,    22

    Me.Width  = 416
    Me.Height = T + 62   ' T = bottom of buttons + title bar / border allowance
End Sub

Private Sub AddCtl(sProgID As String, sName As String, _
                   dL As Single, dT As Single, dW As Single, dH As Single)
    Dim oC As Object
    Set oC = Me.Controls.Add(sProgID, sName, True)
    oC.Left   = dL
    oC.Top    = dT
    oC.Width  = dW
    oC.Height = dH
End Sub

' ---- 2. Set all non-geometry properties ------------------------------------

Private Sub InitControlProps()
    Me.Caption = "Job Application verwalten"

    Me.Controls("lblMailHdr").Caption    = "Ausgewaehlte Mail"
    Me.Controls("lblMailHdr").Font.Bold  = True
    Me.Controls("lblMailFrom").Caption    = "(Von: -)"
    Me.Controls("lblMailSubject").Caption = "(Betreff: -)"
    Me.Controls("lblSep1").BackColor      = RGB(128, 128, 128)
    Me.Controls("lblSep2").BackColor      = RGB(128, 128, 128)

    Me.Controls("optNeu").Caption      = "Neue Bewerbung"
    Me.Controls("optNeu").Value        = True
    Me.Controls("optBestehend").Caption = "Bestehende:"
    Me.Controls("cboJobApplications").Visible = False

    Me.Controls("lblFirma").Caption           = "Firma *"
    Me.Controls("lblPosition").Caption        = "Position *"
    Me.Controls("lblAnsprechpartner").Caption = "Ansprechpartner"
    Me.Controls("lblAnzeigeLink").Caption     = "Anzeigen-Link"
    Me.Controls("lblAnzeigeText").Caption     = "Anzeigentext"
    Me.Controls("txtAnzeigeText").MultiLine   = True
    Me.Controls("txtAnzeigeText").ScrollBars  = 2

    Me.Controls("lblStatus").Caption = "Status"
    Me.Controls("cboStatus").AddItem "geplant"
    Me.Controls("cboStatus").AddItem "gesendet"
    Me.Controls("cboStatus").AddItem "aktiv"
    Me.Controls("cboStatus").AddItem "archiviert"
    Me.Controls("cboStatus").ListIndex = 0

    Me.Controls("lblVorgang").Caption      = "Vorgang / Notiz"
    Me.Controls("lblNotizen").Caption      = "Notizen"
    Me.Controls("txtNotizen").MultiLine    = True
    Me.Controls("txtNotizen").ScrollBars   = 2
    Me.Controls("lblHistorie").Caption     = "Historie"
    Me.Controls("txtHistorie").MultiLine   = True
    Me.Controls("txtHistorie").ScrollBars  = 2
    Me.Controls("txtHistorie").Locked      = True
    Me.Controls("txtHistorie").BackColor   = RGB(240, 240, 240)

    Me.Controls("btnZuordnen").Caption  = "Mail zuordnen"
    Me.Controls("btnSpeichern").Caption = "Nur speichern"
    Me.Controls("btnClose").Caption     = "Schliessen"
    Me.Controls("btnSettings").Caption  = "Settings"
End Sub

' ---- 3. Wire event sinks ---------------------------------------------------

Private Sub WireEvents()
    Set m_Sinks = New Collection

    Dim aBtns As Variant
    aBtns = Array("btnZuordnen", "btnSpeichern", "btnClose", "btnSettings")
    Dim i As Integer
    For i = 0 To UBound(aBtns)
        Dim oBtn As clsBtnSink
        Set oBtn = New clsBtnSink
        Set oBtn.Btn  = Me.Controls(aBtns(i))
        Set oBtn.Form = Me
        m_Sinks.Add oBtn
    Next i

    Dim aOpts As Variant
    aOpts = Array("optNeu", "optBestehend")
    For i = 0 To UBound(aOpts)
        Dim oOpt As clsOptSink
        Set oOpt = New clsOptSink
        Set oOpt.Opt  = Me.Controls(aOpts(i))
        Set oOpt.Form = Me
        m_Sinks.Add oOpt
    Next i

    Dim oCbo As clsCboSink
    Set oCbo = New clsCboSink
    Set oCbo.Cbo  = Me.Controls("cboJobApplications")
    Set oCbo.Form = Me
    m_Sinks.Add oCbo
End Sub

' ---- 4. Load data and pre-fill form ----------------------------------------

Private Sub LoadData()
    Set m_List = New JobApplicationList
    m_List.Load
    PopulateJobCombo

    Set m_Mail = GetCurrentMail()
    If Not m_Mail Is Nothing Then
        Me.Controls("lblMailFrom").Caption    = "Von: " & m_Mail.SenderName & _
                                                "  <" & m_Mail.SenderEmailAddress & ">"
        Me.Controls("lblMailSubject").Caption = "Betreff: " & m_Mail.Subject

        Dim sCats As String
        sCats = m_Mail.Categories
        Dim oApp As JobApplication
        For Each oApp In m_List.GetAll()
            If InStr(sCats, oApp.Tag) > 0 Then
                Set m_App = oApp
                SelectExistingApp oApp.Tag
                Exit For
            End If
        Next oApp
    Else
        Me.Controls("lblMailFrom").Caption    = "(Keine Mail ausgewaehlt)"
        Me.Controls("lblMailSubject").Caption = ""
        Me.Controls("btnZuordnen").Enabled    = False
    End If
End Sub

' =============================================================================
'  Public event dispatchers  (called by sink objects)
' =============================================================================

Public Sub OnButtonClick(sName As String)
    Select Case sName
        Case "btnZuordnen":  ZuordnenAction
        Case "btnSpeichern": SpeichernAction
        Case "btnClose":     Me.Hide
        Case "btnSettings":  Settings.OpenSettingsFile
    End Select
End Sub

Public Sub OnOptionClick(sName As String)
    Select Case sName
        Case "optNeu":       NeuSelected
        Case "optBestehend": BestehendSelected
    End Select
End Sub

Public Sub OnComboChange(sName As String)
    If sName = "cboJobApplications" Then LoadSelectedApp
End Sub

' =============================================================================
'  Button actions
' =============================================================================

Private Sub ZuordnenAction()
    If Not ValidateInput() Then Exit Sub
    Dim oApp As JobApplication
    If Me.Controls("optNeu").Value Then
        Set oApp = BuildNewApp()
        oApp.Create
    Else
        Set oApp = m_App
        If oApp Is Nothing Then
            MsgBox "Bitte eine Bewerbung auswaehlen.", vbExclamation : Exit Sub
        End If
        ApplyEditsTo oApp
    End If
    oApp.AddMail m_Mail, Me.Controls("cboStatus").Text, Me.Controls("txtVorgang").Text
    MsgBox "Mail zugeordnet: " & oApp.Tag, vbInformation, "JobHunt"
    Me.Hide
End Sub

Private Sub SpeichernAction()
    If Not ValidateInput() Then Exit Sub
    Dim oApp As JobApplication
    If Me.Controls("optNeu").Value Then
        Set oApp = BuildNewApp()
        oApp.Create
    Else
        Set oApp = m_App
        If oApp Is Nothing Then
            MsgBox "Bitte eine Bewerbung auswaehlen.", vbExclamation : Exit Sub
        End If
        ApplyEditsTo oApp
        oApp.UpdateAndSave Me.Controls("cboStatus").Text, Me.Controls("txtVorgang").Text
    End If
    MsgBox "Bewerbung gespeichert.", vbInformation, "JobHunt"
    Me.Hide
End Sub

' =============================================================================
'  Option-button actions
' =============================================================================

Private Sub NeuSelected()
    Me.Controls("cboJobApplications").Visible = False
    Me.Controls("txtFirma").Enabled    = True
    Me.Controls("txtPosition").Enabled = True
    ClearFields
    Set m_App = Nothing
End Sub

Private Sub BestehendSelected()
    Me.Controls("cboJobApplications").Visible = True
    If Me.Controls("cboJobApplications").ListCount > 0 Then
        Me.Controls("cboJobApplications").ListIndex = 0
        LoadSelectedApp
    End If
End Sub

' =============================================================================
'  Private helpers
' =============================================================================

Private Function BuildNewApp() As JobApplication
    Dim o As JobApplication
    Set o = New JobApplication
    o.Init Me.Controls("txtFirma").Text, Me.Controls("txtPosition").Text
    ApplyEditsTo o
    Set BuildNewApp = o
End Function

Private Sub ApplyEditsTo(oApp As JobApplication)
    oApp.Ansprechpartner = Me.Controls("txtAnsprechpartner").Text
    oApp.Anzeige_Link    = Me.Controls("txtAnzeigeLink").Text
    oApp.Anzeige_Text    = Me.Controls("txtAnzeigeText").Text
    oApp.Notizen         = Me.Controls("txtNotizen").Text
End Sub

Private Sub PopulateJobCombo()
    Me.Controls("cboJobApplications").Clear
    Dim oApp As JobApplication
    For Each oApp In m_List.GetAll()
        Me.Controls("cboJobApplications").AddItem _
            oApp.ID & " - " & oApp.Firma & " / " & oApp.Position & " [" & oApp.Status & "]"
    Next oApp
End Sub

Private Sub LoadSelectedApp()
    If Me.Controls("cboJobApplications").ListIndex < 0 Then Exit Sub
    Dim sID As String
    sID = Trim(Split(Me.Controls("cboJobApplications").Text, " - ")(0))
    Set m_App = m_List.GetByID(sID)
    If Not m_App Is Nothing Then FillFields m_App
End Sub

Private Sub SelectExistingApp(sTag As String)
    Me.Controls("optBestehend").Value         = True
    Me.Controls("cboJobApplications").Visible = True
    Dim i As Long
    For i = 0 To Me.Controls("cboJobApplications").ListCount - 1
        Dim sID As String
        sID = Trim(Split(Me.Controls("cboJobApplications").List(i), " - ")(0))
        Dim o As JobApplication
        Set o = m_List.GetByID(sID)
        If Not o Is Nothing Then
            If o.Tag = sTag Then
                Me.Controls("cboJobApplications").ListIndex = i
                Exit For
            End If
        End If
    Next i
    If Not m_App Is Nothing Then FillFields m_App
End Sub

Private Sub FillFields(oApp As JobApplication)
    Me.Controls("txtFirma").Text           = oApp.Firma
    Me.Controls("txtPosition").Text        = oApp.Position
    Me.Controls("txtAnsprechpartner").Text = oApp.Ansprechpartner
    Me.Controls("txtAnzeigeLink").Text     = oApp.Anzeige_Link
    Me.Controls("txtAnzeigeText").Text     = oApp.Anzeige_Text
    Me.Controls("txtNotizen").Text         = oApp.Notizen
    Me.Controls("txtHistorie").Text        = oApp.Historie
    Dim i As Long
    For i = 0 To Me.Controls("cboStatus").ListCount - 1
        If Me.Controls("cboStatus").List(i) = oApp.Status Then
            Me.Controls("cboStatus").ListIndex = i : Exit For
        End If
    Next i
End Sub

Private Sub ClearFields()
    Me.Controls("txtFirma").Text           = ""
    Me.Controls("txtPosition").Text        = ""
    Me.Controls("txtAnsprechpartner").Text = ""
    Me.Controls("txtAnzeigeLink").Text     = ""
    Me.Controls("txtAnzeigeText").Text     = ""
    Me.Controls("txtVorgang").Text         = ""
    Me.Controls("txtNotizen").Text         = ""
    Me.Controls("txtHistorie").Text        = ""
    Me.Controls("cboStatus").ListIndex     = 0
End Sub

Private Function ValidateInput() As Boolean
    If Me.Controls("optNeu").Value Then
        If Len(Trim(Me.Controls("txtFirma").Text)) = 0 Then
            MsgBox "Bitte Firma eingeben.", vbExclamation
            Me.Controls("txtFirma").SetFocus : Exit Function
        End If
        If Len(Trim(Me.Controls("txtPosition").Text)) = 0 Then
            MsgBox "Bitte Position eingeben.", vbExclamation
            Me.Controls("txtPosition").SetFocus : Exit Function
        End If
    End If
    ValidateInput = True
End Function

Private Function GetCurrentMail() As Outlook.MailItem
    On Error Resume Next
    Dim oI As Outlook.Inspector
    Set oI = Application.ActiveInspector
    If Not oI Is Nothing Then
        If oI.CurrentItem.Class = olMail Then
            Set GetCurrentMail = oI.CurrentItem : Exit Function
        End If
    End If
    Dim oE As Outlook.Explorer
    Set oE = Application.ActiveExplorer
    If Not oE Is Nothing Then
        If oE.Selection.Count > 0 Then
            If oE.Selection(1).Class = olMail Then
                Set GetCurrentMail = oE.Selection(1)
            End If
        End If
    End If
    On Error GoTo 0
End Function
