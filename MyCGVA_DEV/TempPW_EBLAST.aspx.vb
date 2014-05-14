Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class TempPW_EBLAST
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
    'Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")
    Private connStr As String = "Data Source=whsql-v20.prod.mesa1.secureserver.net;Initial Catalog=cgva;User Id=cgva;Password=Sideout123;"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

            'PWGenerator.PasswordGenerator.CreatePasswords()

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()
            SQL = "SELECT p.FIRST_NAME,p.EMAIL,l.USERNAME,l.TEMP_PASSWORD " _
                & "FROM db_accessadmin.[PERSON_TBL] p " _
                & "LEFT JOIN USER_LOGIN_TBL l " _
                & "ON p.PERSON_ID = l.PERSON_ID " _
                & "WHERE LOGICAL_DELETE_IND = 'N'  and LASTDIG_IND = 'N' " _
                & "AND SUPPRESS_EMAIL_IND = 'N' AND RTRIM(EMAIL) <> '' " _
                & "AND p.PERSON_ID NOT IN (17,19,96,97,265,323,478)"
            '& "AND TEMP_PASSWORD <> ''"
            'Response.Write(SQL)
            'Response.End()
            rs = New SqlCommand(SQL, cn)
            rsReply = rs.ExecuteReader

            While rsReply.Read

                'email the person the welcome letter
                Dim vEmailFrom As String = "cgva@cgva.org"
                Dim emailTo As String = rsReply("EMAIL")
                Dim emailCC As String = ""
                Dim emailBCC As String = "jcrossin@aol.com"
                Dim emailSubject As String = "Welcome to MyCGVA!"
                Dim emailBody As String = "<div>Dear " & rsReply("FIRST_NAME") & ", " _
                                        & "<br /><br />" _
                                        & "The CGVA website team is rolling out a brand new online experience called MyCGVA! Initially, MyCGVA will provide a secure area on the CGVA website where you can:" _
                                        & "<br /><br />" _
                                        & "-Register/pay for CGVA leagues and tournaments" _
                                        & "<br />" _
                                        & "-Set up preferences" _
                                        & "<br />" _
                                        & "-Manage your contact information" _
                                        & "<br />" _
                                        & "-Reset your password and your security questions" _
                                        & "<br /><br />" _
                                        & "We have created an ID and a temporary password for you, which you must change the first time you logon to MyCGVA. " _
                                        & "<br /><br />" _
                                        & "User Name: " & rsReply("USERNAME") _
                                        & "<br />" _
                                        & "Temporary Password: " & rsReply("TEMP_PASSWORD") _
                                        & "<br /><br />" _
                                        & "Please use the link below to go to the login page, change your password and begin using MyCGVA! " _
                                        & "<br /><br />" _
                                        & "<a href='http://www.cgva.org/MyCGVA/Login.aspx'>http://www.cgva.org/MyCGVA/Login.aspx</a>" _
                                        & "<br /><br />" _
                                        & "Note that information that we have on file for you may be out-dated or incorrect.  Feel free to go into the 'My Profile' area of the site to change your information." _
                                       & "<br /><br />" _
                                        & "Please add cgva@cgva.org and support@cgva.org to your address book so you receive all future e-blasts about CGVA and any messages related to your account. " _
                                        & "<br /><br />" _
                                        & "If you are not interested in having an ID on the CGVA website, please send an email to support@cgva.org and we’ll deactivate your account." _
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
                    Me.messageLabel.Text += "<font class='cfontSuccess10'>SUCCESS:" + rsReply("EMAIL") + "<br /></font>"
                Catch ex As Exception
                    Me.messageLabel.Text += "<font class='cfontError10'>FAILURE:" + rsReply("EMAIL") + "</font><br />"
                    'Me.messageLabel.Text = "<font class=""cfontError10"">Sorry, there was an issue sending you the welcome email. Please contact CGVA at support@cgva.org.</font>"
                    Exit Sub
                End Try


            End While

            'Me.messageLabel.Text = "<font class='cfontSuccess10'>Success.</font>"

        End If

    End Sub

End Class

