' ASP.NET VB

Imports System.Net
Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient


Partial Class IPNTEST
    Inherits System.Web.UI.Page

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim strResponse As String = "VERIFIED"
        Dim testText As String = "97~on~on~on~Y~N"


        Dim personID As String = "0"
        Dim aryTextFile() As String

        Dim vEmailFrom As String = "cgva@cgva.org"
        Dim emailTo As String = "jcrossin@aol.com"
        Dim emailCC As String = ""
        Dim emailBCC As String = "support@cgva.org"
        Dim emailSubject As String = ""
        Dim emailBody As String = ""

        Dim em As New MailMessage
        Dim smtp As New SmtpClient
        smtp.Host = EMAIL_SERVER

        If strResponse = "VERIFIED" Then
            'check the payment_status is Completed
            'check that txn_id has not been previously processed
            'check that receiver_email is your Primary PayPal email
            'check that payment_amount/payment_currency are correct
            'process payment
            emailSubject = "VERIFIED - CGVA 2011 SPRING LEAGUE REGISTRATION"
            Try
                aryTextFile = testText.Split("~")

                Try
                    personID = aryTextFile(0)
                Catch ex As Exception
                    personID = "0"
                End Try

                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = personID.ToString + " " + emailSubject
                em.Body = testText
                em.IsBodyHtml = True

                smtp.Send(em)
            Catch ex As Exception
                Exit Sub
            End Try

            'insert a record into the appropriate tables
            Dim cn As SqlConnection
            Dim rs As SqlCommand
            'Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            Dim League As String
            Dim LeagueFee As String = "80"
            Dim LeagueNotes As String = "PAYPAL"

            Dim verifyInfo As String
            Dim verifyWaiverInfo As String

            Dim CAPTAIN_IND_League As String
            Dim Q5 As String

            Try
                '           For Each item In Request.Form
                '            Dim fieldname As String
                '            Dim fieldValue As String
                '            fieldname = item
                '            fieldValue = Request.Form(item)
                '            Response.Write("Form field name:" & fieldname & "Form field value:" & fieldValue & "<BR>")
                '            Next

                personID = aryTextFile(0)
                Session("PERSON_ID") = personID

                League = aryTextFile(1)

                verifyInfo = aryTextFile(2)
                If verifyInfo = "on" Then
                    verifyInfo = "Y"
                Else
                    verifyInfo = "N"
                End If

                verifyWaiverInfo = aryTextFile(3)
                If verifyWaiverInfo = "on" Then
                    verifyWaiverInfo = "Y"
                Else
                    verifyWaiverInfo = "N"
                End If

                CAPTAIN_IND_League = aryTextFile(4)
                Q5 = aryTextFile(5)

                'rating requested?
                If Q5 <> "N" Then
                    SQL = "INSERT INTO RATING_REQUEST_TBL(PERSON_ID) " _
                    & "SELECT '" & personID & "'"
                    rs = New SqlCommand(SQL, cn)
                    rs.ExecuteNonQuery()
                End If

                'registration record(s)
                If League = "on" Then
                    SQL = "If not exists(select PERSON_ID from REGISTRATION_TBL where PERSON_ID='" & personID & "' and EVENT_CD='2011SL' and REGISTRATION_IND='Y') " _
                    & "INSERT INTO REGISTRATION_TBL(" _
                     & "EVENT_CD, " _
                     & "PERSON_ID, " _
                     & "[DATE], " _
                     & "WAIVER_IND, " _
                     & "OPEN_PLAY_IND, " _
                     & "PLAY_UPDOWN_IND, " _
                     & "REGISTRATION_IND, " _
                     & "DOLLARS_COLLECTED, " _
                     & "DOLLARS_OFF_COUPON, " _
                     & "CHECK_AMT_COLLECTED, " _
                     & "CHECK_NUM, " _
                     & "NOTES, " _
                     & "CAPTAIN_INTEREST_IND, " _
                     & "VOLUNTEER_INTEREST_IND)" _
                     & " VALUES(" _
                     & "'2011SL', " _
                     & "'" & personID & "', " _
                     & "getdate(), " _
                     & "'" & verifyWaiverInfo & "', " _
                     & "'N', " _
                     & "'N', " _
                     & "'Y', " _
                     & "'" & LeagueFee & "', " _
                     & "'', " _
                     & "'', " _
                     & "'', " _
                     & "'" & LeagueNotes & "', " _
                     & "'" & CAPTAIN_IND_League & "', " _
                     & "'N')"
                    rs = New SqlCommand(SQL, cn)
                    rs.ExecuteNonQuery()
                End If

                messageLabel.Text = SQL
            Catch ex As Exception
                messageLabel.Text = ex.StackTrace
            Finally
                cn.Close()
                cn = Nothing
            End Try
        ElseIf strResponse = "INVALID" Then
            'log for manual investigation
            emailSubject = "INVALID - CGVA 2011 SPRING LEAGUE REGISTRATION"

            Try
                'aryTextFile = Request.Form("custom").Split("~")
                aryTextFile = testText.Split("~")
                Try
                    personID = aryTextFile(0)
                Catch ex As Exception
                    personID = "0"
                End Try

                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = personID.ToString + " " + emailSubject
                em.Body = Request.Form("custom").ToString
                em.IsBodyHtml = True

                smtp.Send(em)
            Catch ex As Exception
                Exit Sub
            End Try
        Else
            'Response wasn't VERIFIED or INVALID, log for manual investigation
            emailSubject = "Response not VERIFIED or INVALID - CGVA 2011 SPRING LEAGUE REGISTRATION"

            Try
                'aryTextFile = Request.Form("custom").Split("~")
                aryTextFile = testText.Split("~")
                Try
                    personID = aryTextFile(0)
                Catch ex As Exception
                    personID = "0"
                End Try

                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = personID.ToString + " " + emailSubject
                'em.Body = Request.Form("custom").ToString
                em.IsBodyHtml = True

                smtp.Send(em)
            Catch ex As Exception
                Exit Sub
            End Try
        End If

        messageLabel.Text += "?"
    End Sub
End Class


