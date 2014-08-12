' ASP.NET VB

Imports System.Net
Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient


Partial Class IPNHandler
    Inherits System.Web.UI.Page

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Post back to either sandbox or live
        'TEST
        'Dim strSandbox As String = "https://www.sandbox.paypal.com/cgi-bin/webscr"
        'Dim req As HttpWebRequest = CType(WebRequest.Create(strSandbox), HttpWebRequest)
        'PROD
        Dim strLive As String = "https://www.paypal.com/cgi-bin/webscr"
        Dim req As HttpWebRequest = CType(WebRequest.Create(strLive), HttpWebRequest)


        'Set values for the request back
        req.Method = "POST"
        req.ContentType = "application/x-www-form-urlencoded"
        Dim Param() As Byte = Request.BinaryRead(HttpContext.Current.Request.ContentLength)
        Dim strRequest As String = Encoding.ASCII.GetString(Param)
        strRequest = strRequest + "&cmd=_notify-validate"
        req.ContentLength = strRequest.Length

        'for proxy
        'Dim proxy As New WebProxy(New System.Uri("http://url:port#"))
        'req.Proxy = proxy

        'Send the request to PayPal and get the response
        Dim streamOut As StreamWriter = New StreamWriter(req.GetRequestStream(), Encoding.ASCII)
        streamOut.Write(strRequest)
        streamOut.Close()
        Dim streamIn As StreamReader = New StreamReader(req.GetResponse().GetResponseStream())
        Dim strResponse As String = streamIn.ReadToEnd()
        streamIn.Close()

        Dim personID As String = "0"
        Dim aryTextFile() As String

        Dim vEmailFrom As String = "cgva@cgva.org"
        Dim emailTo As String = "jcrossin11@gmail.com"
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
            emailSubject = "VERIFIED - CGVA REGISTRATION"
            Try
                aryTextFile = Request.Form("custom").Split("~")

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

            'insert a record into the appropriate tables
            Dim cn As SqlConnection
            Dim rs As SqlCommand
            'Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            Dim EventCode As String
            Dim LeagueFee As String = "0"
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

                'aryTextFile = Request.QueryString("custom").Split("~")

                personID = aryTextFile(0)

                EventCode = aryTextFile(1) '8/3/2012 now holds the value of the event code
                LeagueFee = aryTextFile(2)

                'verifyInfo = aryTextFile(3)
                'If verifyInfo = "on" Then
                verifyInfo = "Y"
                'Else
                'verifyInfo = "N"
                'End If

                'verifyWaiverInfo = aryTextFile(4)
                'If verifyWaiverInfo = "on" Then
                verifyWaiverInfo = "Y"
                'Else
                'verifyWaiverInfo = "N"
                'End If

                'CAPTAIN_IND_League = aryTextFile(5)
                CAPTAIN_IND_League = "Y"
                'Q5 = aryTextFile(6)
                Q5 = "N"

                'rating requested?
                If Q5 <> "N" Then
                    SQL = "INSERT INTO RATING_REQUEST_TBL(PERSON_ID) " _
                    & "SELECT '" & personID & "'"
                    rs = New SqlCommand(SQL, cn)
                    rs.ExecuteNonQuery()
                End If

                'registration record(s)
                'If EventCode = "on" Then
                SQL = "If not exists(select PERSON_ID from REGISTRATION_TBL where PERSON_ID='" & personID & "' and EVENT_CD='" & EventCode & "' and REGISTRATION_IND='Y') " _
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
                 & "'" & EventCode & "', " _
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
                'End If

            Catch ex As Exception

            Finally
                cn.Close()
                cn = Nothing
            End Try
        ElseIf strResponse = "INVALID" Then
            'log for manual investigation
            emailSubject = "INVALID - CGVA REGISTRATION"

            Try
                aryTextFile = Request.Form("custom").Split("~")

                Try
                    personID = aryTextFile(0)
                    Session("PERSON_ID") = personID
                Catch ex As Exception
                    personID = "0"
                    Session("PERSON_ID") = personID
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
            emailSubject = "Response not VERIFIED or INVALID - CGVA REGISTRATION"
            Try
                aryTextFile = Request.Form("custom").Split("~")

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
        End If
    End Sub
End Class

