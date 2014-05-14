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
        Dim strSandbox As String = "https://www.sandbox.paypal.com/cgi-bin/webscr"
        Dim strLive As String = "https://www.paypal.com/cgi-bin/webscr"
        'TEST
        'Dim req As HttpWebRequest = CType(WebRequest.Create(strSandbox), HttpWebRequest)
        'PROD
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

        Dim teamName As String = "0"
        Dim division As String = "0"
        Dim totalFee As String = "0"
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

        Try
            'teamName + "_" + division + "_" + totalFee
            aryTextFile = Request.Form("custom").Split("~")
            teamName = aryTextFile(0)
            division = aryTextFile(1)
            totalFee = aryTextFile(2)
        Catch ex As Exception
            teamName = "XXX"
        End Try

        If strResponse = "VERIFIED" Then
            'check the payment_status is Completed
            'check that txn_id has not been previously processed
            'check that receiver_email is your Primary PayPal email
            'check that payment_amount/payment_currency are correct
            'process payment
            emailSubject = "VERIFIED - " & Request.Form("custom").ToString & " CGVA 2014 LAST DIG REGISTRATION"
            Try
                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = teamName + " " + emailSubject
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

            Try

                'registration record(s)
                SQL = "If not exists(select ID from LASTDIG_TEAM_TBL " _
                & "where ARCHIVED_DTG IS NULL AND " _
                & "ACTIVE = 'Y' " _
                & "AND DIVISION = '" & division & "' " _
                & "AND TEAM_NAME = '" & teamName & "') " _
                & "INSERT INTO LASTDIG_TEAM_TBL(TEAM_NAME,DIVISION,FEE_PAID, ACTIVE) " _
                            & "SELECT '" & teamName & "', " _
                            & "'" & division & "', " _
                            & "'" & totalFee & "', " _
                            & "'Y' "

                rs = New SqlCommand(SQL, cn)
                rs.ExecuteNonQuery()

            Catch ex As Exception

            Finally
                cn.Close()
                cn = Nothing
            End Try

        ElseIf strResponse = "INVALID" Then
            'log for manual investigation
            emailSubject = "INVALID - " & Request.Form("custom").ToString & " CGVA 2014 LAST DIG REGISTRATION"

            Try

                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = teamName + " " + emailSubject
                em.Body = Request.Form("custom").ToString
                em.IsBodyHtml = True

                smtp.Send(em)
            Catch ex As Exception
                Exit Sub
            End Try
        Else
            'Response wasn't VERIFIED or INVALID, log for manual investigation
            emailSubject = "Response not VERIFIED or INVALID - " & Request.Form("custom").ToString & " CGVA 2014 LAST DIG REGISTRATION"

            Try
                em.From = New MailAddress(vEmailFrom)
                em.To.Add(New MailAddress(emailTo))
                'em.Bcc.Add(New MailAddress(emailBCC))
                em.Subject = teamName + " " + emailSubject
                em.Body = Request.Form("custom").ToString
                em.IsBodyHtml = True

                smtp.Send(em)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub
End Class

