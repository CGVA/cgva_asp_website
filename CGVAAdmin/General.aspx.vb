Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Net.Mail

Partial Class General
    Inherits System.Web.UI.Page

    Public vEmailFrom As String = ""
    Private const string SERVER = "relay-hosting.secureserver.net"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

            'txtMsgLabel.Text = ""
            'txtMsgLabel.Visible = False

            Session("USER_ID") = Request("USER_ID")
            Session("MGMT") = Request("MGMT")
            Session("MGMT_DATA") = Request("MGMT_DATA")
            Session("CUSTOM_DISTRO") = Request("CUSTOM_DISTRO")
            Session("SAVEDISTRO") = Request("SAVEDISTRO")
            Session("emailFrom") = Request("emailFrom")
            Session("emailTo") = Request("emailTo")
            Session("emailSubject") = Request("emailSubject")
            Session("emailCC") = Request("emailCC")
            Session("emailBCC") = Request("emailBCC")
            Session("ADD_ADDRESS") = Request("ADD_ADDRESS")
            Session("TEMPLATE") = Request("TEMPLATE")

            Dim connstring As String = System.Configuration.ConfigurationManager.AppSettings("SDATConnectionString")
            Dim cnNav As New SqlConnection
            cnNav.ConnectionString = connstring

            If mnuNav Is Nothing Then

                Exit Sub

            End If



            'mnuNav.Items.Add(New MenuItem("&nbsp;Home | ", "/BRT/Home.asp"))

            'If Session("ROLE_ID") <= 1 Then
            '    mnuNav.Items.Add(New MenuItem("Admin | ", "/BRT/Admin.asp"))
            'End If

            'mnuNav.Items.Add(New MenuItem("Manage Distros | ", "/BRT/Distros.asp"))
            'mnuNav.Items.Add(New MenuItem("Reports | ", "/BRT/Reports.asp"))
            'mnuNav.Items.Add(New MenuItem("Open A Ticket | ", "http://ivue.vzbi.com/home/tickets.cfm  target='_blank'"))

            'mnuNav.Items.Add(New MenuItem("Logout", "/logout.asp"))


            If UCase(Request.ServerVariables("SERVER_NAME")) = "LOCALHOST" Then

                Session("USER_ID") = "1"
                Session("TEMPLATE") = "General"
                'Session("MGMT") = "6425,8351,13700"
                Session("CUSTOM_DISTRO") = ""
                Session("ADD_ADDRESS") = "jcrossin@aol.com"
                Session("emailTo") = ""
                Session("emailBCC") = ""
                Session("MGMT_DATA") = "JIM ELDER"

            End If

            Dim TEMPLATE As String = ""
            TEMPLATE = Session("TEMPLATE")

            Dim MGMT As String = ""
            MGMT = Session("MGMT")

            vEmailFrom = "cgva@cgva.org"

            Dim txtSub As TextBox = Me.FORM.FindControl("txtEmailFrom")
            txtSub.Text = vEmailFrom
            'txtSub.Focus()
            txtEmailSubject.Text = " "


            Dim emailTo As String = Session("emailTo")
            Dim emailCC As String = ""
            Dim emailBCC As String = "jcrossin@aol.com"
            Dim emailSubject As String = ""
            Dim emailBody As String = ""

            Dim SQL As String = ""

            'Response.Write("USER_ID:" & Session("USER_ID"))

		If Session("ACCESS") = "" then
			Session("Err") = "Your session has timed out. Please log in again."
			Response.Redirect("Adminindex.asp")
		ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
			Session("Err") = "You do not have access to view the requested page."
			Response.Redirect("Adminindex.asp")
		End If


            'emailTo = ""
            'emailCC = ""
            'emailSubject = ""
            'emailBody = ""

            If Session("emailBCC") = "" And Session("emailTo") = "" Then

                If Session("ADD_ADDRESS") <> "" Then
                    Session("emailBCC") = Session("ADD_ADDRESS") & "; "
                End If

                SQL = "EMAILTEMPLATE_EMAIL @CUSTOM_DISTRO = '" & Replace(Session("CUSTOM_DISTRO"), "'", "''") & "', @MGMT='" & Session("MGMT") & "'"

                ''rw("<!-- SQL: " & SQL & " -->")
                'Response.Write(SQL)
                'Response.End()

                Dim txtBcc As TextBox = Me.FORM.FindControl("txtEmailBcc")
                Dim cn As SqlConnection
                cn = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SDATConnectionString"))
                cn.Open()
                Dim reply As New SqlCommand(SQL, cn)
                Dim rs As SqlDataReader = reply.ExecuteReader
                If rs.HasRows Then

                    For Each record As Common.DbDataRecord In rs
                        'Session("emailBCC") = Session("emailBCC") & record(0) & ";"
                        txtBcc.Text = txtBcc.Text & record(0) & ";"
                    Next

                    txtBcc.Text = txtBcc.Text.Substring(0, txtBcc.Text.Length - 1)
                    'txtBcc.Text = Session("emailBCC")
                    'Response.Write("BCC:" & Session("emailBCC") & ", TO:" & Session("emailTo"))


                End If
                cn.Close()
                cn = Nothing


                'testing - LIMIT SEEMS TO BE 127 total addresses, 128 bombs out
                'Dim tempString As String = "jcrossin@aol.com;"
                'For i As Integer = 0 To 2000
                'tempString = tempString + "john.crossin" + i.ToString + "@verizonbusiness.com;"
                'Next

                'txtBcc.Text = tempString
                ''end testing

                'Else
                'Response.Write("BCC:" & Session("emailBCC") & ", TO:" & Session("emailTo"))
                'Response.End()

            End If

        End If


    End Sub

    Protected Sub mnuNav_MenuItemClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.MenuEventArgs) Handles mnuNav.MenuItemClick
        If e.Item.Value <> "" Then
            Response.Redirect(e.Item.Value)
        End If
    End Sub

    Sub rw(ByVal vString As String)
        Response.Write(vString)
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim sep(6) As Char
        sep(0) = ","
        sep(1) = ":"
        sep(2) = ";"

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim txtAddresses As TextBox = FORM.FindControl("txtEmailTo")
        'Response.Write("ALL TO:" & txtAddresses.Text & "<br />")

        Dim tempString As String
        tempString = txtAddresses.Text.Replace("\s", "")
        tempString = tempString.Replace(" ", "")
        tempString = tempString.Replace(vbCRLF, "")

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''test validity of all emails first
        Dim em_addresses() As String = tempString.Split(sep)
        For Each em_add As String In em_addresses

            Try
                If em_add.Length > 0 Then
                    'Response.Write("TO:" & em_add & "<br />")
                    Dim tempAddress = New MailAddress(Trim(em_add))
                End If
            Catch ex As Exception
                'Response.Write("ERROR")
                'Response.End()
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: unable to add TO email address: " & em_add & ".</b></font>"
                Exit Sub

            End Try

        Next

        'Dim emCC_addresses() As String = e.Values("emailCC").ToString.Split("; ")
        Dim txtCCAddresses As TextBox = FORM.FindControl("txtEmailCC")
        'Response.Write("ALL CC:" & txtCCAddresses.Text & "<br />")

        tempString = ""
        tempString = txtCCAddresses.Text.Replace("\s", "")
        tempString = tempString.Replace(" ", "")
        tempString = tempString.Replace(vbCRLF, "")
        Dim emCC_addresses() As String = tempString.Split(sep)

        For Each em_add As String In emCC_addresses
            Try
                If em_add.Length > 0 Then
                    'Response.Write("CC:" & em_add & "<br />")
                    Dim tempAddress = New MailAddress(Trim(em_add))
                End If
            Catch ex As Exception
                'Response.Write("ERROR")
                'Response.End()
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: unable to add CC email address: " & em_add & ".</b></font>"
                Exit Sub

            End Try
        Next

        'Dim emBCC_addresses() As String = e.Values("emailBCC").ToString.Split("; ")
        Dim txtBCCAddresses As TextBox = FORM.FindControl("txtEmailBcc")
        'Response.Write("ALL BCC:" & txtBCCAddresses.Text & "<br />")
        'sites = values.Split(sep, 4)

        tempString = ""
        tempString = txtBCCAddresses.Text.Replace("\s", "")
        tempString = tempString.Replace(" ", "")
        tempString = tempString.Replace(vbCRLF, "")
        'Response.Write(tempString)
        'Response.End()
        Dim emBCC_addresses() As String = tempString.Split(sep)

        For Each em_add As String In emBCC_addresses
            Try
                If em_add.Length > 0 Then
                    'Response.Write("BCC:" & em_add & "<br />")
                    Dim tempAddress = New MailAddress(Trim(em_add))
                End If
            Catch ex As Exception
                'Response.Write("ERROR:" & ex.ToString)
                'Response.End()
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: unable to add BCC email address: " & em_add & ".</b></font>"
                Exit Sub

            End Try
        Next

        '' add me to all deliveries
        ''em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        Dim em As New MailMessage

        Dim txtEmailFrom As TextBox = FORM.FindControl("txtEmailFrom")
        em.From = New MailAddress(txtEmailFrom.Text)

        Dim txtSub As TextBox = FORM.FindControl("txtEmailSubject")
        em.Subject = txtSub.Text


        ''attachments
        'getting the file name and storing it in the server path and then attaching it to the 'mail. 
        Dim mailattach As Attachment 'creating new instance of MailAttachment object 
        Dim attachname1 As String 'string which stores the attachment name 
        Dim attachname2 As String 'string which stores the second attachment name 

        attachname1 = FileUpload1.PostedFile.FileName
        attachname2 = FileUpload2.PostedFile.FileName

        If (attachname1 <> "") Then
            attachname1 = attachname1.Substring(attachname1.LastIndexOf("\") + 1)
            'Response.Write(attachname1)
            Try
                FileUpload1.PostedFile.SaveAs(Server.MapPath("docs/") & attachname1)
                mailattach = New Attachment(Server.MapPath("docs/") & attachname1)
                em.Attachments.Add(mailattach)

            Catch ex As Exception
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There is already an attachment stored on the server with the same name. Please rename your attachment, and try again.</b></font>"
                Exit Sub
            End Try
        End If

        'getting the file name and storing it in the server path and then attaching it to the 'mail. 
        If (attachname2 <> "") Then
            attachname2 = attachname2.Substring(attachname2.LastIndexOf("\") + 1)
            'Response.Write(attachname2)
            Try
                FileUpload2.PostedFile.SaveAs(Server.MapPath("docs/") & attachname2)
                mailattach = New Attachment(Server.MapPath("docs/") & attachname2)
                em.Attachments.Add(mailattach)

            Catch ex As Exception
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There is already an attachment stored on the server with the same name. Please rename your attachment, and try again.</b></font>"
                Exit Sub
            End Try

        End If

        'em.Attachments.Add(New Attachment(Server.MapPath("docs/option2.txt")))

        Dim txtBody As Moxiecode.TinyMCE.Web.TextArea = FORM.FindControl("letter")
        'em.Body = e.Values("emailBody")

'        em.Body = "<img src='http://sdautomation.vzbi.com/EmailTemplate/images/VPSLogo.jpg' border='0' /><p />" & txtBody.Value


        em.IsBodyHtml = True
'        Dim smtp As New SmtpClient
'        smtp.Host = "smtp.vzbi.com"

	Dim em As MailMessage = New System.Web.Mail.MailMessage() 
	em.From = "emailaddress@domainname" 
	em.[To] = "emailaddress@domainname" 
	em.Subject = "Test email subject" 
	em.BodyFormat = MailFormat.Html 
	' enumeration 
	em.Priority = MailPriority.High 
	' enumeration 
	em.Body = "Sent at: " & DateTime.Now 
	SmtpMail.SmtpServer = SERVER 
	SmtpMail.Send(em) 
	' free up resources 
	em = Nothing 
 
 	EXIT SUB		'COME BACK TO THIS!
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim counter As Integer = 1


        For Each em_add As String In em_addresses
            'Response.Write("TO:" & em_add & em_add.length & "<br />")
            ''gather sets of 100 email addresses and send 
            If counter Mod 100 > 0 Then

                Try
                    If em_add.Length > 0 Then
                        em.To.Add(New MailAddress(Trim(em_add)))
                        counter = counter + 1
                    End If
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the TO email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                End Try


            ElseIf counter Mod 100 = 0 Then
                ''add the last record to the email

                Try
                    em.To.Add(New MailAddress(Trim(em_add)))

                    ''add me as BCC
                    em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                    smtp.Send(em)
                    em.To.Clear()
                    em.Bcc.Clear()
                    ''reset counter
                    counter = 1
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the TO email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                    'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                    'Response.Redirect("Home.asp")
                End Try

            End If

        Next

        'hit the end of the array, any strays to send?
        If em.To.ToString <> "" Then

            Try
                ''add me as BCC
                em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                smtp.Send(em)
                em.To.Clear()
                em.Bcc.Clear()
                ''reset counter
                counter = 1
            Catch ex As Exception
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the TO email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                Exit Sub
                'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                'Response.Redirect("Home.asp")
            End Try
        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''reset counter for CC email
        counter = 1
        For Each emCC_add As String In emCC_addresses

            ''gather sets of 100 email addresses and send 
            If counter Mod 100 > 0 Then

                Try
                    If emCC_add.Length > 0 Then
                        em.CC.Add(New MailAddress(Trim(emCC_add)))
                        counter = counter + 1
                    End If
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the CC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                End Try


            ElseIf counter Mod 100 = 0 Then
                ''add the last record to the email
                em.CC.Add(New MailAddress(Trim(emCC_add)))

                Try
                    ''add me as BCC
                    em.To.Clear()
                    em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                    smtp.Send(em)
                    em.CC.Clear()
                    em.Bcc.Clear()
                    ''reset counter
                    counter = 1
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the CC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                    'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                    'Response.Redirect("Home.asp")
                End Try

            End If

        Next

        'hit the end of the array, any strays to send?
        If em.CC.ToString <> "" Then

            Try
                ''add me as BCC
                em.To.Add(New MailAddress("VPS-Customer-Service@Verizonbusiness.com"))
                em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                smtp.Send(em)
                em.To.Clear()
                em.CC.Clear()
                em.Bcc.Clear()
                ''reset counter
                counter = 1
            Catch ex As Exception
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the CC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                Exit Sub
                'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                'Response.Redirect("Home.asp")
            End Try

        End If

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''reset counter for BCC email
        counter = 1
        For Each emBCC_add As String In emBCC_addresses
            'Response.Write("BCC:" & emBCC_add & "<br />")
            ''gather sets of 100 email addresses and send 
            If counter Mod 100 > 0 Then

                Try
                    If emBCC_add.Length > 0 Then
                        em.Bcc.Add(New MailAddress(Trim(emBCC_add)))
                        counter = counter + 1
                    End If
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the BCC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                End Try


            ElseIf counter Mod 100 = 0 Then
                ''add the last record to the email
                em.Bcc.Add(New MailAddress(Trim(emBCC_add)))

                Try
                    ''add me as BCC

                    em.To.Add(New MailAddress("VPS-Customer-Service@Verizonbusiness.com"))
                    em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                    smtp.Send(em)
                    em.To.Clear()
                    em.Bcc.Clear()
                    ''reset counter
                    counter = 1
                Catch ex As Exception
                    txtMsgLabel.Visible = True
                    txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the BCC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                    Exit Sub
                    'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                    'Response.Redirect("Home.asp")
                End Try

            End If

        Next

        'hit the end of the array, any strays to send?
        If em.Bcc.ToString <> "" Then

            Try
                ''add me as BCC
                em.To.Add(New MailAddress("VPS-Customer-Service@Verizonbusiness.com"))
                em.Bcc.Add(New MailAddress("jcrossin@aol.com"))
                smtp.Send(em)
                em.To.Clear()
                em.Bcc.Clear()
                ''reset counter
                counter = 1
            Catch ex As Exception
                txtMsgLabel.Visible = True
                txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error sending the BCC email. Please contact SWS using the 'Open a Ticket' link on the navigation bar. " & ex.ToString & "</b></font>"
                Exit Sub
                'Session("Err") = "<font class='cfontError10'><b>Error: There was an error sending the email. Please contact SWS using the 'Open a Ticket' link on the navigation bar.</b></font>"
                'Response.Redirect("Home.asp")
            End Try
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try

'            Dim SQL As String = ""

'            SQL = "INSERT INTO T_EMAILTEMPLATE_HISTORY(USER_ID,TEMPLATE,SUBJECT) VALUES(" & _
'            "'" & Session("USER_ID") & "', " & _
'            "'" & Session("TEMPLATE") & "', " & _
'            "'" & Replace(txtEmailSubject.Text, "'", "''") & "')"

'            Dim cn As SqlConnection
'            cn = New SqlConnection(System.Configuration.ConfigurationManager.AppSettings("SDATConnectionString"))
'            cn.Open()
'            Dim reply As New SqlCommand(SQL, cn)
'            reply.ExecuteNonQuery()
'            cn.Close()
'            cn = Nothing

        Catch ex As Exception
            txtMsgLabel.Visible = True
            txtMsgLabel.Text = "<font face='Arial' color='#990000'><b>Error: There was an error inserting a record into the email history table. " & ex.ToString & "</b></font>"
            Exit Sub
        End Try

        txtMsgLabel.Visible = True
        txtMsgLabel.Text = "<font face='Arial' color='#009900'><b>Your email has been sent successfully.</b></font>"


    End Sub
End Class