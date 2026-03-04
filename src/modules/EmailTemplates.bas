Attribute VB_Name = "EmailTemplates"
'==============================================================================
' Module     : EmailTemplates
' Description: Pre-built email templates for common job-hunting scenarios.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Opens a new email pre-filled with a job application cover letter template.
'------------------------------------------------------------------------------
Public Sub NewApplicationEmail()
    Dim oMail As Outlook.MailItem
    Set oMail = Application.CreateItem(olMailItem)

    oMail.Subject = "Application for [Position] at [Company]"
    oMail.Body = _
        "Dear Hiring Team," & vbNewLine & vbNewLine & _
        "I am writing to express my interest in the [Position] role at [Company]." & vbNewLine & vbNewLine & _
        "Please find my CV attached. I am confident that my skills in [Skill1] and [Skill2] make me a strong candidate." & vbNewLine & vbNewLine & _
        "I look forward to the opportunity to discuss how I can contribute to your team." & vbNewLine & vbNewLine & _
        "Kind regards," & vbNewLine & _
        Application.Session.CurrentUser.Name

    oMail.Display
End Sub

'------------------------------------------------------------------------------
' Opens a new follow-up email after an interview.
'------------------------------------------------------------------------------
Public Sub NewFollowUpEmail()
    Dim oMail As Outlook.MailItem
    Set oMail = Application.CreateItem(olMailItem)

    oMail.Subject = "Thank You – [Position] Interview"
    oMail.Body = _
        "Dear [Name]," & vbNewLine & vbNewLine & _
        "Thank you for taking the time to meet with me [today / on DATE] to discuss the [Position] role." & vbNewLine & vbNewLine & _
        "I enjoyed learning more about [Company] and am very excited about the opportunity." & vbNewLine & vbNewLine & _
        "Please do not hesitate to contact me if you require any further information." & vbNewLine & vbNewLine & _
        "Kind regards," & vbNewLine & _
        Application.Session.CurrentUser.Name

    oMail.Display
End Sub

'------------------------------------------------------------------------------
' Opens a new status-enquiry email.
'------------------------------------------------------------------------------
Public Sub NewStatusEnquiryEmail()
    Dim oMail As Outlook.MailItem
    Set oMail = Application.CreateItem(olMailItem)

    oMail.Subject = "Application Status Enquiry – [Position]"
    oMail.Body = _
        "Dear [Name]," & vbNewLine & vbNewLine & _
        "I am writing to enquire about the status of my application for the [Position] role submitted on [DATE]." & vbNewLine & vbNewLine & _
        "I remain very interested in the position and would welcome any update you are able to provide." & vbNewLine & vbNewLine & _
        "Kind regards," & vbNewLine & _
        Application.Session.CurrentUser.Name

    oMail.Display
End Sub
