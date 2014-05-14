Imports System.Data.SqlClient
Imports System.Net.Mail


Partial Class LoginHelp
    Inherits System.Web.UI.Page
    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"

    Private EMAIL As String
    Private USERNAME As String
    Private PASSWORD As String
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        'check validity of email address'
        EMAIL = Trim(Me.EMAIL_ADDRESS.Text)

        If EMAIL = "" Then
            successLabel.Text = "<font class=""cfontError10"">Please enter a vaild email address.</font>"
        Else
            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            SQL = "SELECT ul.USERNAME,ul.PASSWORD,UL.TEMP_PASSWORD " _
            & "FROM USER_LOGIN_TBL ul " _
            & "LEFT JOIN db_accessadmin.[PERSON_TBL] p " _
            & "ON ul.PERSON_ID = p.PERSON_ID " _
            & "WHERE UPPER(p.EMAIL) = '" & UCase(EMAIL) & "'"

            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            'set rs = cn.(SQL)

            If rsreply.Read Then
                'information found = proceed
                USERNAME = rsreply("USERNAME")

                If rsreply("TEMP_PASSWORD") <> "" Then
                    PASSWORD = rsreply("TEMP_PASSWORD")
                Else
                    PASSWORD = rsreply("PASSWORD")
                End If

                'send email
                Dim vEmailFrom As String = "cgva@cgva.org"
                Dim emailTo As String = EMAIL
                Dim emailCC As String = ""
                Dim emailBCC As String = "support@cgva.org"
                Dim emailSubject As String = "MyCGVA User Name/Password Request"
                Dim emailBody As String = "Your MyCGVA user name is: " & USERNAME & " <br />Your password is: " & Me.PASSWORD
                Dim em As New MailMessage
                Dim smtp As New SmtpClient
                smtp.Host = EMAIL_SERVER

                Try
                    em.From = New MailAddress(vEmailFrom)
                    em.To.Add(New MailAddress(emailTo))
                    em.Bcc.Add(New MailAddress(emailBCC))
                    em.Subject = emailSubject
                    em.Body = emailBody
                    em.IsBodyHtml = True

                    smtp.Send(em)
                Catch ex As Exception
                    Me.successLabel.Text = "<font class=""cfontError10"">Sorry, there was an issue with the email address you entered. Please check the address and try again, or contact CGVA at support@cgva.org.</font>"
                    Exit Sub
                End Try

                Me.successLabel.Text = "Your user name and password were successfully sent to you.<br />Please check your email for this information, then click <a href='Login.aspx'>here</a> to access MyCGVA."
            Else
                'email address not found
                Me.successLabel.Text = "<font class=""cfontError10"">Sorry, the email address entered was not found.</font>"
            End If

            rsreply.Close()
            cn.Close()


        End If
    End Sub

End Class
