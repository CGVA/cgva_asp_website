Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class NewRegistration
    Inherits System.Web.UI.Page

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"

    Private userName, userName1, userName2, userName3 As String
    Private FNAME As String
    Private LNAME As String
    Private strEmail As String
    Private personID As Integer
    Private firstContactID As Integer
    Private strGender As String
    Private strTshirt As String
    Private strPhone1 As String
    Private strPhone2 As String
    Private strPassword1 As String
    Private strPassword2 As String
    Private securityQuestion As Integer
    Private securityQuestionText As String
    Private securityAnswer As String

    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        FNAME = Trim(Me.FIRST_NAME.Text)
        LNAME = Trim(Me.LAST_NAME.Text)
        strEmail = Trim(Me.EMAIL.Text)
        firstContactID = Me.FIRST_CONTACT_ID.SelectedValue
        strGender = Me.GENDER.SelectedValue
        strTshirt = Me.TSHIRT.SelectedValue
        strPhone1 = Trim(Me.PRIMARY_PHONE_NUM.Text)
        strPhone2 = Trim(Me.PHONE2.Text)
        strPassword1 = Trim(Me.PASSWORD1.Text)
        strPassword2 = Trim(Me.PASSWORD2.Text)
        securityQuestion = Me.SECURITY_QUESTION.SelectedValue
        securityQuestionText = Me.SECURITY_QUESTION.SelectedItem.Text
        securityAnswer = Trim(Me.SECURITY_ANSWER.Text)

        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsReply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr


        If FNAME = "" Then
            Me.messageLabel.Text = "Please enter your first name."
            Exit Sub
        ElseIf LNAME = "" Then
            Me.messageLabel.Text = "Please enter your last name."
            Exit Sub
        ElseIf strEmail = "" Then
            Me.messageLabel.Text = "Please enter your email address."
            Exit Sub
        ElseIf firstContactID = 0 Then
            Me.messageLabel.Text = "Please enter how you first heard about CGVA."
            Exit Sub
        ElseIf strPassword1 = "" Then
            Me.messageLabel.Text = "Please enter a password."
            Exit Sub
        ElseIf strPassword1.Length < 6 Then
            Me.messageLabel.Text = "Please enter a password between 6 and 20 characters in length."
            Exit Sub
        ElseIf strPassword2 = "" Then
            Me.messageLabel.Text = "Please verify the password."
            Exit Sub
        ElseIf strPassword1 <> strPassword2 Then
            Me.messageLabel.Text = "Please make sure the password and the re-typed password are identical."
            Exit Sub
        ElseIf securityQuestion = 0 Then
            Me.messageLabel.Text = "Please select a security question."
            Exit Sub
        ElseIf securityAnswer = "" Then
            Me.messageLabel.Text = "Please supply an answer to the selected security question."
            Exit Sub
        Else
            'make sure the supplied email address doesnt already exist
            cn.Open()
            SQL = "SELECT COUNT(*) as 'COUNT' FROM db_accessadmin.[PERSON_TBL] WHERE UPPER(EMAIL) = '" & UCase(Me.strEmail) & "'"
            'Response.Write(SQL)
            'Response.End()

            rs = New SqlCommand(SQL, cn)
            rsReply = rs.ExecuteReader

            While rsReply.Read

                If rsReply(0) <> 0 Then
                    Me.messageLabel.Text = "Sorry, this email address already exists in our records."
                    cn.Close()
                    cn.Dispose()
                    Exit Sub
                End If

            End While
            rsReply.Close()

        End If

        'Add this validated record to the database
        userName1 = UCase(Me.FNAME.Substring(0, 1) & Me.LNAME)
        userName2 = UCase(Me.FNAME.Substring(0, 1) & Me.LNAME & "1")
        userName3 = UCase(Me.FNAME.Substring(0, 1) & Me.LNAME & "2")

        'check to see if the default username is already taken for this person
        'if so, add a digit to the end of the name
        SQL = "SELECT COUNT(*) FROM USER_LOGIN_TBL WHERE UPPER(USERNAME) ='" & userName1 & "' " _
        & "SELECT COUNT(*) FROM USER_LOGIN_TBL WHERE UPPER(USERNAME) ='" & userName2 & "' " _
        & "SELECT COUNT(*) FROM USER_LOGIN_TBL WHERE UPPER(USERNAME) ='" & userName3 & "' "

        rs = New SqlCommand(SQL, cn)
        rsReply = rs.ExecuteReader

        If rsReply.Read Then

            If rsReply(0) = 0 Then
                Me.userName = userName1
            Else
                rsReply.NextResult()

                If rsReply.Read Then

                    If rsReply(0) = 0 Then
                        Me.userName = userName2
                    Else
                        rsReply.NextResult()

                        If rsReply.Read Then

                            If rsReply(0) = 0 Then
                                Me.userName = userName3
                            Else
                                Me.messageLabel.Text = "Sorry, we are unable to create your login credentials at this time. Please contact us at support@cgva.org for assistance."
                                cn.Close()
                                cn.Dispose()
                                Exit Sub
                            End If

                        End If

                    End If

                End If
            End If

            rsReply.Close()

            SQL = "INSERT INTO db_accessadmin.[PERSON_TBL](" _
            & "FIRST_NAME, " _
            & "LAST_NAME, " _
            & "EMAIL, " _
            & "GENDER, " _
            & "PRIMARY_PHONE_NUM, " _
            & "[2ND_PHONE_NUM], " _
            & "FIRST_CONTACT_ID, " _
            & "TSHIRT_SIZE, " _
            & "SUPPRESS_EMAIL_IND,SUPPRESS_SNAIL_MAIL_IND,SUPPRESS_LAST_NAME_IND) " _
            & "SELECT '" & Me.FNAME & "', " _
            & "'" & Me.LNAME & "', " _
            & "'" & Me.strEmail & "', " _
            & "'" & Me.strGender & "', " _
            & "'" & Me.strPhone1 & "', " _
            & "'" & Me.strPhone2 & "', " _
            & "'" & Me.firstContactID.ToString & "', " _
            & "'" & Me.strTshirt & "', " _
            & "'N','N','N'"

            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()
            'get this persons PERSON_ID now to enter a record into USER_LOGIN_TBL
            SQL = "SELECT PERSON_ID FROM db_accessadmin.[PERSON_TBL] " _
                & "WHERE EMAIL='" & Me.strEmail & "'"
            rs = New SqlCommand(SQL, cn)
            rsReply = rs.ExecuteReader

            If rsReply.Read Then
                Me.personID = rsReply(0)
            End If
            rsReply.Close()

            'enter record into USER_LOGIN_TBL
            SQL = "INSERT INTO USER_LOGIN_TBL(" _
            & "PERSON_ID, " _
            & "USERNAME, " _
            & "PASSWORD, " _
            & "SECURITY_QUESTION, " _
            & "ANSWER) " _
            & "SELECT '" & Me.personID.ToString & "', " _
            & "'" & Me.userName & "', " _
            & "'" & Me.strPassword1 & "', " _
            & "'" & Me.securityQuestion.ToString & "', " _
            & "'" & Me.securityAnswer & "'"

            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()
            cn.Close()

            'email the person the welcome letter
            Dim vEmailFrom As String = "cgva@cgva.org"
            Dim emailTo As String = Me.strEmail
            Dim emailCC As String = ""
            Dim emailBCC As String = "support@cgva.org"
            Dim emailSubject As String = "Welcome to MyCGVA!"
            Dim emailBody As String = "<div>Dear " & Me.FNAME & ", " _
                                    & "<br /><br />" _
                                    & "Thanks for registering with MyCGVA! We are glad you have chosen to be a part of the CGVA community and we hope you enjoy using our online services. " _
                                    & "<br /><br />" _
                                    & "Please add cgva@cgva.org and support@cgva.org to your address book so you receive all future e-blasts about CGVA and any messages related to your account. " _
                                    & "<br /><br />" _
                                    & "First Name: " & Me.FNAME & "<br />" _
                                    & "Last Name: " & Me.LNAME & "<br />" _
                                    & "Email: " & Me.strEmail & "<br />" _
                                    & "Username: " & Me.userName & "<br />" _
                                    & "Password: " & Me.strPassword1 & "<br />" _
                                    & "Security Question: " & Me.securityQuestionText & "<br />" _
                                    & "Answer: " & Me.securityAnswer & "<br />" _
                                    & "<br /><br />" _
                                    & "If you have not registered on our site or the information is sent in error, please send an email right away to support@cgva.org so we can fix the problem." _
                                    & "<br /><br />" _
                                    & "Sincerely," _
                                    & "<br /><br />" _
                                    & "The CGVA Website Team" _
                                    & "<br /><br />" _
                                    & "For General info/questions about CGVA send emails to cgva@cgva.org " _
                                    & "<br />" _
                                    & "For Support issues/problems with the CGVA Website send emails to support@cgva.org </div>"

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
                Me.messageLabel.Text = "<font class=""cfontError10"">Sorry, there was an issue sending you the welcome email. Please contact CGVA at support@cgva.org.</font>"
                Exit Sub
            End Try

            Me.messageLabel.Text = "<font class='cfontSuccess10'>Your user name and password were successfully sent to you.<br />Please check your email for this information, then click <a href='Login.aspx'>here</a> to access MyCGVA.</font>"

        End If

    End Sub

Protected Sub PASSWORD1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PASSWORD1.TextChanged

End Sub
End Class
