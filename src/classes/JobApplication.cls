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
' Description: Represents a single job application derived from an email.
'==============================================================================
Option Explicit

Private m_Company    As String
Private m_Position   As String
Private m_AppliedOn  As Date
Private m_Status     As String
Private m_EntryID    As String   ' Outlook MailItem EntryID

' ---- Properties --------------------------------------------------------------

Public Property Get Company() As String
    Company = m_Company
End Property
Public Property Let Company(sValue As String)
    m_Company = sValue
End Property

Public Property Get Position() As String
    Position = m_Position
End Property
Public Property Let Position(sValue As String)
    m_Position = sValue
End Property

Public Property Get AppliedOn() As Date
    AppliedOn = m_AppliedOn
End Property
Public Property Let AppliedOn(dValue As Date)
    m_AppliedOn = dValue
End Property

Public Property Get Status() As String
    Status = m_Status
End Property
Public Property Let Status(sValue As String)
    m_Status = sValue
End Property

Public Property Get EntryID() As String
    EntryID = m_EntryID
End Property
Public Property Let EntryID(sValue As String)
    m_EntryID = sValue
End Property

' ---- Methods -----------------------------------------------------------------

'------------------------------------------------------------------------------
' Factory: populates this instance from an Outlook MailItem.
'------------------------------------------------------------------------------
Public Sub LoadFromMail(oMail As Outlook.MailItem)
    Me.EntryID  = oMail.EntryID
    Me.AppliedOn = oMail.SentOn

    ' Attempt to extract Company / Position from the subject line.
    ' Expected pattern: "Application for <Position> at <Company>"
    Dim subject As String
    subject = oMail.Subject

    Dim parts() As String
    If InStr(subject, " at ") > 0 Then
        parts = Split(subject, " at ", 2)
        Me.Company = Trim(parts(1))

        Dim leftPart As String
        leftPart = parts(0)
        If InStr(leftPart, " for ") > 0 Then
            Me.Position = Trim(Split(leftPart, " for ", 2)(1))
        Else
            Me.Position = Trim(leftPart)
        End If
    Else
        Me.Company  = "Unknown"
        Me.Position = subject
    End If

    Me.Status = "Pending"
End Sub

'------------------------------------------------------------------------------
' Returns a formatted one-line summary.
'------------------------------------------------------------------------------
Public Function Summary() As String
    Summary = Format(m_AppliedOn, "YYYY-MM-DD") & _
              "  |  " & m_Company & _
              "  |  " & m_Position & _
              "  |  " & m_Status
End Function
