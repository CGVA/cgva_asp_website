<%@ Language=VBScript %>
<!-- #include virtual = "/includes/SDATconnection.inc" -->
<!-- #include virtual="/includes/rw.asp" -->

<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	dim exists
	dim ID, role, active, lastName, firstName, loginID, password
	dim roleString,activeString,lastNameString,firstNameString,loginIDString,pwString
	dim IDarray, rolearray, activearray, lastNamearray, firstNamearray, loginIDarray, passwordarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("index.asp")
	ElseIf Instr(Session("ACCESS"),"ADMIN") < 0 and Instr(Session("ACCESS"),"EDIT") < 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("index.asp")
	End If

	If Request("submitChoice") = "Add User" then

		role = Request("role")
		loginID = Request("loginID")
		search = Request("search")
		active = Request("active")

		'check to see if the login ID entered exists'
		checkLoginID()

		if exists <> "Y" then
			sql = "INSERT INTO T_SDAT_LOGIN("_
					& "ROLE_ID,"_
					& "LOGIN_ID,"_
					& "ACTIVE_FLAG)"_
					& "VALUES("_
					& role & ", "_
					& "'" & loginID & "', "_
					& "'" & active & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		end if

		''get userID value for this new user
		sql = "SELECT USER_ID FROM T_SDAT_LOGIN "_
			& "WHERE UPPER(LOGIN_ID) = '" & UCase(loginID) & "'"

		set rs = cn.Execute(sql)
		USER_ID = rs(0)

		''''''''''' ADD RECORDS TO T_SDAT_ACCESS TABLE '''''
		''check to see if this user already has access to BRT
		sql = "SELECT * FROM T_SDAT_ACCESS WHERE USER_ID='" & USER_ID & "' "_
			& "AND APP_ID='8'"

		set rs = cn.Execute(sql)

		if rs.EOF then

			sql = "INSERT INTO T_SDAT_ACCESS("_
					& "USER_ID,"_
					& "APP_ID)"_
					& "VALUES("_
					& "'" & USER_ID & "', "_
					& "'8')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)

			call closeRSCNConnection()
			Session("admin") = "insert"
			Response.Redirect("Personnel.asp")

		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Personnel.asp")
		end if


	ElseIf Request("submitChoice") = "Modify User(s)" then

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		role = Request("role")
		rolearray = split(role,", ")

		loginID = Request("loginID")
		loginIDarray = split(loginID,", ")

		search = Request("search")
		searcharray = split(search,", ")

		active = Request("active")
		activearray = split(active,", ")

		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)

			checkLoginID()

			if exists <> "Y" then
				sql = "UPDATE T_SDAT_LOGIN SET "_
						& "ROLE_ID = " & rolearray(i) & ", "_
						& "LOGIN_ID = '" & loginIDarray(i) & "', "_
						& "ACTIVE_FLAG = '" & activearray(i) & "' "_
						& "WHERE USER_ID = " & IDarray(i)
				''rw(sql)
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & loginIDarray(i)
			end if

		Next

		''Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Personnel.asp?notupdated=" & notUpdated)

	Else
		rw(Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Personnel.asp")
	End If


'******************************************'

Sub checkLoginID()
	If Request("submitChoice") = "Modify User(s)" then
		sql = "SELECT LOGIN_ID FROM T_SDAT_LOGIN WHERE UPPER(LOGIN_ID)='" & UCase(loginIDarray(i)) & "' "
		sql = sql & "AND USER_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT LOGIN_ID FROM T_SDAT_LOGIN WHERE UPPER(LOGIN_ID)='" & UCase(loginID) & "' "

	End If

	''rw(sql)
	''Response.End

	set rs = cn.Execute(sql)

	If not rs.EOF then
		exists = "Y"
	Else
		exists = "N"
	End If

End Sub

'******************************************'

%>