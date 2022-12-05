Imports System.Net.Mail
Public Class SendMail

    Private _smtpServer As String
    Public Property SmtpServer() As String
        Get
            Return _smtpServer
        End Get
        Set(ByVal value As String)
            _smtpServer = value
        End Set
    End Property


    Private _emailTo As ArrayList
    Public Property EmailTo() As ArrayList
        Get
            Return _emailTo
        End Get
        Set(ByVal value As ArrayList)
            _emailTo = value
        End Set
    End Property

    Private _emailFrom As String
    Public Property EmailFrom() As String
        Get
            Return _emailFrom
        End Get
        Set(ByVal value As String)
            _emailFrom = value
        End Set
    End Property

    Private _message As String
    Public Property Message() As String
        Get
            Return _message
        End Get
        Set(ByVal value As String)
            _message = value
        End Set
    End Property

    Private _subject As String
    Public Property Subject() As String
        Get
            Return _subject
        End Get
        Set(ByVal value As String)
            _subject = value
        End Set
    End Property

    Private _body As String
    Public Property Body() As String
        Get
            Return _body
        End Get
        Set(ByVal value As String)
            _body = value
        End Set
    End Property

    Private _isBodyHtml As Boolean
    Public Property IsBodyHtml() As Boolean
        Get
            Return _isBodyHtml
        End Get
        Set(ByVal value As Boolean)
            _isBodyHtml = value
        End Set
    End Property

    Private _attachment As ArrayList
    Public Property Attachment() As ArrayList
        Get
            Return _attachment
        End Get
        Set(ByVal value As ArrayList)
            _attachment = value
        End Set
    End Property


    Public Sub Send()
        Using Smtp_Svr = New SmtpClient(SmtpServer)
            'Using Smtp_Svr = New SmtpClient("edi-mail.pki.com.ph")
            Dim mail = New MailMessage
            Smtp_Svr.DeliveryMethod = SmtpDeliveryMethod.Network
            With mail
                '.From = New MailAddress("donotreply@pki.lip")
                .From = New MailAddress(EmailFrom)
                For Each email As String In EmailTo
                    .To.Add(email)
                Next
                .IsBodyHtml = IsBodyHtml
                .Subject = Subject

                .Body = Body

                If Not IsNothing(Attachment) Then
                    For Each fileName As String In Attachment
                        .Attachments.Add(New Attachment(fileName))
                    Next
                End If
            End With
            Smtp_Svr.Send(mail)
        End Using
    End Sub


End Class
