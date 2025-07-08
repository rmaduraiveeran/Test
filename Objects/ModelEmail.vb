Public Class ModelEmail
    Public Property SmtpClient As String
    Public Property MailSender As String
    Public Property MailRecipients As List(Of String)
    Public Property MailSubject As String
    Public Property MailBody As String
    Public Property MailAttachments As List(Of String)
End Class
