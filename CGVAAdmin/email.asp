<!-- #include virtual="/incs/rw.asp" -->
<%
	''On Error Resume Next

	dim rsEmailData, rsEmailCols,rsEmailRows, SQL,URL,serverChoice,BCC,emailTo,emailCC
	dim rs,cn
	dim rsData, rsCols,rsRows
	dim templateName,pathToTemplate,docName,pathToExcel
	dim xlApp, xlWb, xlWs
	dim Counter

	serverChoice="dev"
	''serverChoice="prod"

	Set cn = CreateObject("ADODB.Connection")

	If serverChoice="dev" Then
		dbServer = "whsql-v19.prod.mesa1.secureserver.net"
		dbName = "cgva_dev"

		''DEV
		dbUsername = "cgva_dev"
		dbPassword = "Sideout456"

		cn.ConnectionString = "Driver={SQL Server}; " _
						& "Server=" & dbServer & "; " _
						& "Database=" & dbName & "; " _
						& "User ID=" & dbUsername & "; " _
						& "Password=" & dbPassword & ";"
	Else
		dbServer = "whsql-v19.prod.mesa1.secureserver.net"
		dbName = "cgva_dev"

		''PROD
		''dbUsername = "cgva"
		''dbPassword = "Sideout123"

		cn.ConnectionString = "Driver={SQL Server}; " _
						& "Server=" & dbServer & "; " _
						& "Database=" & dbName & "; " _
						& "User ID=" & dbUsername & "; " _
						& "Password=" & dbPassword & ";"
	End If

	Set rs=CreateObject("ADODB.Recordset")
	cn.CommandTimeout = 240
	cn.Open

''	SQL = "get_autoemail"
''	set rs = cn.Execute(SQL)

''	if not rs.EOF then
''		rsEmailData = rs.GetRows()
''		rsEmailCols = Ubound(rsEmailData,1)
''		rsEmailRows = Ubound(rsEmailData,2)

''		''LOOP THROUGH ALL THIS''
''		For Counter = 0 to rsEmailRows
''			call SendEmailOut(Counter)
''		Next
''	end if

	call SendEmailOut(0)
	set rs = Nothing
	set cn = Nothing
	set fs = Nothing


	'''''''''''''''''''''''''''

	Sub SendEmailOut(MDCounter)
		Dim objMail
		Set objMail = Server.CreateObject("CDONTS.NewMail")
		objMail.From = "cgva@cgva.org"
		objMail.Subject = "CGVA ID#: "
		objMail.To = "jcrossin@aol.com"
		objMail.Body = "Hi There"
		objMail.Send

		Set objMail = Nothing
''		rw("here1")
''		Response.End

		Response.Redirect("admin.asp")
		'<- auto-redirection You must always do this with CDONTS.
		Response.End


'		CDO.From = "cgva@cgva.org"
'		EMAIL_BODY = ""
'		EMAIL_BODY = EMAIL_BODY & "Hi there"
'
'		''CDO.To = rsEmailData(x,Counter)
'		BCC = "jcrossin@aol.com"
'
'		''user email
'	''	If not IsNull(rsEmailData(8,Counter)) Then
'	''		emailTo = rsEmailData(8,Counter)
'	''	End If
'
'		CDO.Subject = "CGVA ID#: "
'
'		CDO.To = emailTo
'		CDO.Cc = emailCC
'		CDO.Bcc = BCC
'		CDO.HTMLBody = EMAIL_BODY
'
'		CDO.Send
'
'		''reset variables
'		emailTo = ""
'		emailCC = ""
'
'		Set CDO = Nothing
'		Set CDOCon = Nothing
	End Sub

	'''''''''''''''''''''''''''
%>