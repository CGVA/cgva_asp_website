<% 
	Dim cn, rs

	Set cn = Server.CreateObject("ADODB.Connection")
	cn.CommandTimeout = 0

	Set rs = Server.CreateObject("ADODB.Recordset")

''for each x in Request.ServerVariables
''  response.write(x & ":" & Request.ServerVariables(x) & "<br />")
''next


	
	
	''If InStr(UCase(Request.ServerVariables("PATH_INFO")),"/CGVAADMINDEV") > 0 or Instr(Request.ServerVariables("HTTP_REFERER"),"index.asp") > 0 Then
	If InStr(UCase(Request.ServerVariables("PATH_INFO")),"/CGVAADMINDEV") > 0 Then
		''DEV
		'dbServer = "whsql-v19.prod.mesa1.secureserver.net"
		dbServer = "cgvadev.db.2095697.hostedresource.com"
		'dbName = "cgva_dev"
		dbName = "cgvadev"
		'dbUsername = "cgva_dev"
		dbUsername = "cgvadev"
		dbPassword = "Sideout456"
	Else
		''PROD
		'dbServer = "whsql-v20.prod.mesa1.secureserver.net"
		dbServer = "cgva.db.2095697.hostedresource.com"
		dbName = "cgva"
		dbUsername = "cgva"
		dbPassword = "Sideout123"
	End If

	''SQL Server version: 2005 (Direct access allowed) something like this http://www.connectionstrings.com/sql-server-2005/

	cn.ConnectionString = "Driver={SQL Server}; " _
						& "Server=" & dbServer & "; " _
						& "Database=" & dbName & "; " _
						& "User ID=" & dbUsername & "; " _
						& "Password=" & dbPassword & ";"


	cn.Open


	Sub CloseCNConnection()
		cn.Close
		set cn = Nothing
	End Sub

	Sub CloseRSCNConnection()
		rs.Close
		set rs = Nothing
		cn.Close
		set cn = Nothing
	End Sub

 %>