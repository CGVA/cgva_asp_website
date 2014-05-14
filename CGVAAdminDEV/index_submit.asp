<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	USERNAME = Ucase(Replace(Request("USERNAME"),"'","''"))
	PASSWORD = Ucase(Replace(Request("PASSWORD"),"'","''"))

	sql = "SELECT ul.PERSON_ID, IsNull(pa.ROLE_TYPE,'') "_
		& "FROM USER_LOGIN_TBL ul "_
		& "LEFT JOIN PERSON_ACCESS_TBL pa "_
		& "ON ul.PERSON_ID = pa.PERSON_ID "_
		& "WHERE UPPER(USERNAME) = '" & USERNAME & "' "_
		& "AND UPPER(PASSWORD) = '" & PASSWORD & "'"

	''rw(sql)
	''Response.End
	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsData = rs.GetRows
		rsRows = UBound(rsData,2)

		Session("PERSON_ID") = rsData(0,0)
		''get access levels
		For i = 0 to rsRows
			Session("ACCESS") = Session("ACCESS") & rsData(1,i) & ","
		Next

		''enter a successful log entry
		SQL = "INSERT INTO USER_LOGIN_LOG_ENTRY_TBL("_
			& "ENTRY_USERNAME, "_
			& "ENTRY_PASSWORD, "_
			& "ENTRY_DATETIME, "_
			& "ENTRY_SUCCESS_IND) "_
			& "SELECT '" &	USERNAME & "', "_
			& "'" &	PASSWORD & "', "_
			& "getdate(), "_
			& "'Y'"
		cn.Execute(SQL)

		closeRSCNConnection()
		Response.Redirect("admin.asp")
	else
		''enter an unsuccessful log entry
		SQL = "INSERT INTO USER_LOGIN_LOG_ENTRY_TBL("_
			& "ENTRY_USERNAME, "_
			& "ENTRY_PASSWORD, "_
			& "ENTRY_DATETIME, "_
			& "ENTRY_SUCCESS_IND) "_
			& "SELECT '" &	USERNAME & "', "_
			& "'" &	PASSWORD & "', "_
			& "getdate(), "_
			& "'N'"
		cn.Execute(SQL)

		Session("Err") = "The User ID/Password combination was not found. Please re-enter the information."
		closeRSCNConnection()
		Response.Redirect("index.asp")
	end if

%>
