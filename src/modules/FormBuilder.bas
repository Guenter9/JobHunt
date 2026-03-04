Attribute VB_Name = "FormBuilder"
'==============================================================================
' Module     : FormBuilder
' Description: Programmatically creates the frmJobApplication UserForm.
'
'  Run once (or after each code change) to build/rebuild the form:
'    Call FormBuilder.RebuildForm  from the VBA Immediate window
'
'  Key design decision:
'    Controls.Add() in the VBE Designer returns a limited proxy object.
'    In Outlook VBA only geometry (Left/Top/Width/Height) and Name can be set
'    reliably at design time.  ALL other properties (Caption, MultiLine,
'    ScrollBars, Locked, BackColor, Value, Visible, Font.*) must be set at
'    runtime.  They are set in the generated InitControls() sub which is
'    called first from UserForm_Initialize.
'
'  Requirements:
'    Outlook > File > Options > Trust Center > Trust Center Settings >
'    Macro Settings > "Trust access to the VBA project object model" enabled.
'==============================================================================
Option Explicit

' ---- Public entry point ------------------------------------------------------

Public Sub RebuildForm()
    Dim oVBP As Object
    Dim oVBC As Object
    Dim oDes As Object

    On Error GoTo ErrHandler

    Set oVBP = Application.VBE.ActiveVBProject

    ' Remove old form if present
    On Error Resume Next
    oVBP.VBComponents.Remove oVBP.VBComponents("frmJobApplication")
    On Error GoTo ErrHandler

    ' Add new UserForm  (vbext_ct_MSForm = 3)
    Set oVBC = oVBP.VBComponents.Add(3)
    oVBC.Name = "frmJobApplication"

    Set oDes = oVBC.Designer
    oDes.Width  = 400
    oDes.Height = 600

    ' ------------------------------------------------------------------
    ' Add controls - ONLY geometry here; all other properties are set
    ' at runtime in the generated InitControls() sub.
    ' ------------------------------------------------------------------
    Dim L1 As Single: L1 = 6       ' label left edge
    Dim L2 As Single: L2 = 114     ' input left edge
    Dim W1 As Single: W1 = 102     ' label width
    Dim W2 As Single: W2 = 268     ' input/combo width
    Dim H  As Single: H  = 16      ' single-line height
    Dim HM As Single: HM = 54      ' multiline height
    Dim T  As Single: T  = 4       ' running vertical position

    ' Mail info labels
    AC oDes, "Forms.Label.1",         "lblMailHdr",         L1, T,      W1+W2+6, H  : T = T + H + 2
    AC oDes, "Forms.Label.1",         "lblMailFrom",        L1, T,      W1+W2+6, H  : T = T + H + 2
    AC oDes, "Forms.Label.1",         "lblMailSubject",     L1, T,      W1+W2+6, H  : T = T + H + 4
    ' Separator 1
    AC oDes, "Forms.Label.1",         "lblSep1",            L1, T,      W1+W2+6, 2  : T = T + 5
    ' Neu / Bestehend row
    AC oDes, "Forms.OptionButton.1",  "optNeu",             L1, T,      102,     H
    AC oDes, "Forms.OptionButton.1",  "optBestehend",       114, T,     100,     H
    AC oDes, "Forms.ComboBox.1",      "cboJobApplications", 220, T,     162,     H  : T = T + H + 4
    ' Separator 2
    AC oDes, "Forms.Label.1",         "lblSep2",            L1, T,      W1+W2+6, 2  : T = T + 5
    ' Input fields
    AC oDes, "Forms.Label.1",         "lblFirma",           L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtFirma",           L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblPosition",        L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtPosition",        L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblAnsprechpartner", L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtAnsprechpartner", L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblAnzeigeLink",     L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtAnzeigeLink",     L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblAnzeigeText",     L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtAnzeigeText",     L2, T,      W2,      HM : T = T + HM + 3
    AC oDes, "Forms.Label.1",         "lblStatus",          L1, T,      W1,      H
    AC oDes, "Forms.ComboBox.1",      "cboStatus",          L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblVorgang",         L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtVorgang",         L2, T,      W2,      H  : T = T + H + 3
    AC oDes, "Forms.Label.1",         "lblNotizen",         L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtNotizen",         L2, T,      W2,      HM : T = T + HM + 3
    AC oDes, "Forms.Label.1",         "lblHistorie",        L1, T,      W1,      H
    AC oDes, "Forms.TextBox.1",       "txtHistorie",        L2, T,      W2,      HM : T = T + HM + 6
    ' Buttons
    AC oDes, "Forms.CommandButton.1", "btnZuordnen",        L1,  T,     110,     22
    AC oDes, "Forms.CommandButton.1", "btnSpeichern",       122, T,     110,     22
    AC oDes, "Forms.CommandButton.1", "btnClose",           238, T,      76,     22
    AC oDes, "Forms.CommandButton.1", "btnSettings",        320, T,      56,     22

    ' Resize form to fit
    oDes.Height = T + 42

    ' Inject all VBA code (including InitControls which sets all properties)
    InjectCode oVBC

    MsgBox "frmJobApplication wurde erfolgreich erstellt.", vbInformation, "FormBuilder"
    Exit Sub

ErrHandler:
    MsgBox "Fehler " & Err.Number & ": " & Err.Description, vbCritical, "FormBuilder"
End Sub

' ---- Only adds control with geometry; NO other properties set here ----------

Private Sub AC(oDes As Object, sProgID As String, sName As String, _
               dL As Single, dT As Single, dW As Single, dH As Single)
    Dim oC As Object
    Set oC = oDes.Controls.Add(sProgID, sName, True)
    oC.Left   = dL
    oC.Top    = dT
    oC.Width  = dW
    oC.Height = dH
End Sub

' ---- Inject all form VBA code -----------------------------------------------

Private Sub InjectCode(oVBC As Object)
    Dim M As Object
    Set M = oVBC.CodeModule
    If M.CountOfLines > 0 Then M.DeleteLines 1, M.CountOfLines

    Dim s As String

    s = s & "Option Explicit" & vbLf
    s = s & "" & vbLf
    s = s & "Private m_Mail As Outlook.MailItem" & vbLf
    s = s & "Private m_App  As JobApplication" & vbLf
    s = s & "Private m_List As JobApplicationList" & vbLf
    s = s & "" & vbLf

    ' ===== InitControls - sets ALL non-geometry properties at runtime =========
    s = s & "Private Sub InitControls()" & vbLf
    s = s & "    Me.Caption = ""Job Application verwalten""" & vbLf
    s = s & "    lblMailHdr.Caption     = ""Ausgewaehlte Mail""" & vbLf
    s = s & "    lblMailHdr.Font.Bold   = True" & vbLf
    s = s & "    lblMailFrom.Caption    = ""(Von: -)""" & vbLf
    s = s & "    lblMailSubject.Caption = ""(Betreff: -)""" & vbLf
    s = s & "    lblSep1.BackColor      = RGB(128,128,128)" & vbLf
    s = s & "    lblSep2.BackColor      = RGB(128,128,128)" & vbLf
    s = s & "    optNeu.Caption         = ""Neue Bewerbung""" & vbLf
    s = s & "    optNeu.Value           = True" & vbLf
    s = s & "    optBestehend.Caption   = ""Bestehende:""" & vbLf
    s = s & "    cboJobApplications.Visible = False" & vbLf
    s = s & "    lblFirma.Caption           = ""Firma *""" & vbLf
    s = s & "    lblPosition.Caption        = ""Position *""" & vbLf
    s = s & "    lblAnsprechpartner.Caption = ""Ansprechpartner""" & vbLf
    s = s & "    lblAnzeigeLink.Caption     = ""Anzeigen-Link""" & vbLf
    s = s & "    lblAnzeigeText.Caption     = ""Anzeigentext""" & vbLf
    s = s & "    txtAnzeigeText.MultiLine   = True" & vbLf
    s = s & "    txtAnzeigeText.ScrollBars  = 2" & vbLf
    s = s & "    lblStatus.Caption          = ""Status""" & vbLf
    s = s & "    cboStatus.AddItem ""geplant""" & vbLf
    s = s & "    cboStatus.AddItem ""gesendet""" & vbLf
    s = s & "    cboStatus.AddItem ""aktiv""" & vbLf
    s = s & "    cboStatus.AddItem ""archiviert""" & vbLf
    s = s & "    cboStatus.ListIndex = 0" & vbLf
    s = s & "    lblVorgang.Caption        = ""Vorgang / Notiz""" & vbLf
    s = s & "    lblNotizen.Caption        = ""Notizen""" & vbLf
    s = s & "    txtNotizen.MultiLine      = True" & vbLf
    s = s & "    txtNotizen.ScrollBars     = 2" & vbLf
    s = s & "    lblHistorie.Caption       = ""Historie""" & vbLf
    s = s & "    txtHistorie.MultiLine     = True" & vbLf
    s = s & "    txtHistorie.ScrollBars    = 2" & vbLf
    s = s & "    txtHistorie.Locked        = True" & vbLf
    s = s & "    txtHistorie.BackColor     = RGB(240,240,240)" & vbLf
    s = s & "    btnZuordnen.Caption  = ""Mail zuordnen""" & vbLf
    s = s & "    btnSpeichern.Caption = ""Nur speichern""" & vbLf
    s = s & "    btnClose.Caption     = ""Schliessen""" & vbLf
    s = s & "    btnSettings.Caption  = ""Settings""" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    ' ===== UserForm_Initialize ===============================================
    s = s & "Private Sub UserForm_Initialize()" & vbLf
    s = s & "    InitControls" & vbLf
    s = s & "" & vbLf
    s = s & "    Set m_List = New JobApplicationList" & vbLf
    s = s & "    m_List.Load" & vbLf
    s = s & "    PopulateJobCombo" & vbLf
    s = s & "" & vbLf
    s = s & "    Set m_Mail = GetCurrentMail()" & vbLf
    s = s & "    If Not m_Mail Is Nothing Then" & vbLf
    s = s & "        lblMailFrom.Caption    = ""Von: "" & m_Mail.SenderName &" & vbLf
    s = s & "                                 ""  <"" & m_Mail.SenderEmailAddress & "">""" & vbLf
    s = s & "        lblMailSubject.Caption = ""Betreff: "" & m_Mail.Subject" & vbLf
    s = s & "        Dim sCats As String" & vbLf
    s = s & "        sCats = m_Mail.Categories" & vbLf
    s = s & "        Dim oApp As JobApplication" & vbLf
    s = s & "        For Each oApp In m_List.GetAll()" & vbLf
    s = s & "            If InStr(sCats, oApp.Tag) > 0 Then" & vbLf
    s = s & "                Set m_App = oApp" & vbLf
    s = s & "                SelectExistingApp oApp.Tag" & vbLf
    s = s & "                Exit For" & vbLf
    s = s & "            End If" & vbLf
    s = s & "        Next oApp" & vbLf
    s = s & "    Else" & vbLf
    s = s & "        lblMailFrom.Caption    = ""(Keine Mail ausgewaehlt)""" & vbLf
    s = s & "        lblMailSubject.Caption = """"" & vbLf
    s = s & "        btnZuordnen.Enabled    = False" & vbLf
    s = s & "    End If" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    ' ===== Option button events ===============================================
    s = s & "Private Sub optNeu_Click()" & vbLf
    s = s & "    cboJobApplications.Visible = False" & vbLf
    s = s & "    txtFirma.Enabled    = True" & vbLf
    s = s & "    txtPosition.Enabled = True" & vbLf
    s = s & "    ClearFields" & vbLf
    s = s & "    Set m_App = Nothing" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub optBestehend_Click()" & vbLf
    s = s & "    cboJobApplications.Visible = True" & vbLf
    s = s & "    If cboJobApplications.ListCount > 0 Then" & vbLf
    s = s & "        cboJobApplications.ListIndex = 0" & vbLf
    s = s & "        LoadSelectedApp" & vbLf
    s = s & "    End If" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub cboJobApplications_Change()" & vbLf
    s = s & "    LoadSelectedApp" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    ' ===== Button events =====================================================
    s = s & "Private Sub btnZuordnen_Click()" & vbLf
    s = s & "    If Not ValidateInput() Then Exit Sub" & vbLf
    s = s & "    Dim oApp As JobApplication" & vbLf
    s = s & "    If optNeu.Value Then" & vbLf
    s = s & "        Set oApp = New JobApplication" & vbLf
    s = s & "        oApp.Init txtFirma.Text, txtPosition.Text" & vbLf
    s = s & "        oApp.Ansprechpartner = txtAnsprechpartner.Text" & vbLf
    s = s & "        oApp.Anzeige_Link    = txtAnzeigeLink.Text" & vbLf
    s = s & "        oApp.Anzeige_Text    = txtAnzeigeText.Text" & vbLf
    s = s & "        oApp.Notizen         = txtNotizen.Text" & vbLf
    s = s & "        oApp.Create" & vbLf
    s = s & "    Else" & vbLf
    s = s & "        Set oApp = m_App" & vbLf
    s = s & "        If oApp Is Nothing Then" & vbLf
    s = s & "            MsgBox ""Bitte eine Bewerbung auswaehlen."", vbExclamation : Exit Sub" & vbLf
    s = s & "        End If" & vbLf
    s = s & "        oApp.Ansprechpartner = txtAnsprechpartner.Text" & vbLf
    s = s & "        oApp.Anzeige_Link    = txtAnzeigeLink.Text" & vbLf
    s = s & "        oApp.Anzeige_Text    = txtAnzeigeText.Text" & vbLf
    s = s & "        oApp.Notizen         = txtNotizen.Text" & vbLf
    s = s & "    End If" & vbLf
    s = s & "    oApp.AddMail m_Mail, cboStatus.Text, txtVorgang.Text" & vbLf
    s = s & "    MsgBox ""Mail zugeordnet: "" & oApp.Tag, vbInformation, ""JobHunt""" & vbLf
    s = s & "    Me.Hide" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub btnSpeichern_Click()" & vbLf
    s = s & "    If Not ValidateInput() Then Exit Sub" & vbLf
    s = s & "    Dim oApp As JobApplication" & vbLf
    s = s & "    If optNeu.Value Then" & vbLf
    s = s & "        Set oApp = New JobApplication" & vbLf
    s = s & "        oApp.Init txtFirma.Text, txtPosition.Text" & vbLf
    s = s & "        oApp.Ansprechpartner = txtAnsprechpartner.Text" & vbLf
    s = s & "        oApp.Anzeige_Link    = txtAnzeigeLink.Text" & vbLf
    s = s & "        oApp.Anzeige_Text    = txtAnzeigeText.Text" & vbLf
    s = s & "        oApp.Notizen         = txtNotizen.Text" & vbLf
    s = s & "        oApp.Create" & vbLf
    s = s & "    Else" & vbLf
    s = s & "        Set oApp = m_App" & vbLf
    s = s & "        If oApp Is Nothing Then" & vbLf
    s = s & "            MsgBox ""Bitte eine Bewerbung auswaehlen."", vbExclamation : Exit Sub" & vbLf
    s = s & "        End If" & vbLf
    s = s & "        oApp.Ansprechpartner = txtAnsprechpartner.Text" & vbLf
    s = s & "        oApp.Anzeige_Link    = txtAnzeigeLink.Text" & vbLf
    s = s & "        oApp.Anzeige_Text    = txtAnzeigeText.Text" & vbLf
    s = s & "        oApp.Notizen         = txtNotizen.Text" & vbLf
    s = s & "        oApp.UpdateAndSave cboStatus.Text, txtVorgang.Text" & vbLf
    s = s & "    End If" & vbLf
    s = s & "    MsgBox ""Bewerbung gespeichert."", vbInformation, ""JobHunt""" & vbLf
    s = s & "    Me.Hide" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub btnClose_Click()" & vbLf
    s = s & "    Me.Hide" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub btnSettings_Click()" & vbLf
    s = s & "    Settings.OpenSettingsFile" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    ' ===== Private helpers ====================================================
    s = s & "Private Sub PopulateJobCombo()" & vbLf
    s = s & "    cboJobApplications.Clear" & vbLf
    s = s & "    Dim oApp As JobApplication" & vbLf
    s = s & "    For Each oApp In m_List.GetAll()" & vbLf
    s = s & "        cboJobApplications.AddItem oApp.ID & "" - "" & oApp.Firma &" & vbLf
    s = s & "            "" / "" & oApp.Position & "" ["" & oApp.Status & ""]""" & vbLf
    s = s & "    Next oApp" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub LoadSelectedApp()" & vbLf
    s = s & "    If cboJobApplications.ListIndex < 0 Then Exit Sub" & vbLf
    s = s & "    Dim sID As String" & vbLf
    s = s & "    sID = Trim(Split(cboJobApplications.Text, "" - "")(0))" & vbLf
    s = s & "    Set m_App = m_List.GetByID(sID)" & vbLf
    s = s & "    If Not m_App Is Nothing Then FillFields m_App" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub SelectExistingApp(sTag As String)" & vbLf
    s = s & "    optBestehend.Value = True" & vbLf
    s = s & "    cboJobApplications.Visible = True" & vbLf
    s = s & "    Dim i As Long" & vbLf
    s = s & "    For i = 0 To cboJobApplications.ListCount - 1" & vbLf
    s = s & "        Dim sID As String" & vbLf
    s = s & "        sID = Trim(Split(cboJobApplications.List(i), "" - "")(0))" & vbLf
    s = s & "        Dim o As JobApplication" & vbLf
    s = s & "        Set o = m_List.GetByID(sID)" & vbLf
    s = s & "        If Not o Is Nothing Then" & vbLf
    s = s & "            If o.Tag = sTag Then cboJobApplications.ListIndex = i : Exit For" & vbLf
    s = s & "        End If" & vbLf
    s = s & "    Next i" & vbLf
    s = s & "    FillFields m_App" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub FillFields(oApp As JobApplication)" & vbLf
    s = s & "    txtFirma.Text           = oApp.Firma" & vbLf
    s = s & "    txtPosition.Text        = oApp.Position" & vbLf
    s = s & "    txtAnsprechpartner.Text = oApp.Ansprechpartner" & vbLf
    s = s & "    txtAnzeigeLink.Text     = oApp.Anzeige_Link" & vbLf
    s = s & "    txtAnzeigeText.Text     = oApp.Anzeige_Text" & vbLf
    s = s & "    txtNotizen.Text         = oApp.Notizen" & vbLf
    s = s & "    txtHistorie.Text        = oApp.Historie" & vbLf
    s = s & "    Dim i As Long" & vbLf
    s = s & "    For i = 0 To cboStatus.ListCount - 1" & vbLf
    s = s & "        If cboStatus.List(i) = oApp.Status Then" & vbLf
    s = s & "            cboStatus.ListIndex = i : Exit For" & vbLf
    s = s & "        End If" & vbLf
    s = s & "    Next i" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Sub ClearFields()" & vbLf
    s = s & "    txtFirma.Text = """" : txtPosition.Text = """" : txtAnsprechpartner.Text = """"" & vbLf
    s = s & "    txtAnzeigeLink.Text = """" : txtAnzeigeText.Text = """" : txtVorgang.Text = """"" & vbLf
    s = s & "    txtNotizen.Text = """" : txtHistorie.Text = """" : cboStatus.ListIndex = 0" & vbLf
    s = s & "End Sub" & vbLf
    s = s & "" & vbLf

    s = s & "Private Function ValidateInput() As Boolean" & vbLf
    s = s & "    If optNeu.Value Then" & vbLf
    s = s & "        If Len(Trim(txtFirma.Text)) = 0 Then" & vbLf
    s = s & "            MsgBox ""Bitte Firma eingeben."", vbExclamation" & vbLf
    s = s & "            txtFirma.SetFocus : Exit Function" & vbLf
    s = s & "        End If" & vbLf
    s = s & "        If Len(Trim(txtPosition.Text)) = 0 Then" & vbLf
    s = s & "            MsgBox ""Bitte Position eingeben."", vbExclamation" & vbLf
    s = s & "            txtPosition.SetFocus : Exit Function" & vbLf
    s = s & "        End If" & vbLf
    s = s & "    End If" & vbLf
    s = s & "    ValidateInput = True" & vbLf
    s = s & "End Function" & vbLf
    s = s & "" & vbLf

    s = s & "Private Function GetCurrentMail() As Outlook.MailItem" & vbLf
    s = s & "    On Error Resume Next" & vbLf
    s = s & "    Dim oI As Outlook.Inspector" & vbLf
    s = s & "    Set oI = Application.ActiveInspector" & vbLf
    s = s & "    If Not oI Is Nothing Then" & vbLf
    s = s & "        If oI.CurrentItem.Class = olMail Then" & vbLf
    s = s & "            Set GetCurrentMail = oI.CurrentItem : Exit Function" & vbLf
    s = s & "        End If" & vbLf
    s = s & "    End If" & vbLf
    s = s & "    Dim oE As Outlook.Explorer" & vbLf
    s = s & "    Set oE = Application.ActiveExplorer" & vbLf
    s = s & "    If Not oE Is Nothing Then" & vbLf
    s = s & "        If oE.Selection.Count > 0 Then" & vbLf
    s = s & "            If oE.Selection(1).Class = olMail Then" & vbLf
    s = s & "                Set GetCurrentMail = oE.Selection(1)" & vbLf
    s = s & "            End If" & vbLf
    s = s & "        End If" & vbLf
    s = s & "    End If" & vbLf
    s = s & "    On Error GoTo 0" & vbLf
    s = s & "End Function" & vbLf

    M.AddFromString s
End Sub
