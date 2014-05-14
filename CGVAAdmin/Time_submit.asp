<%@ Language=VBScript %>
<!-- #include virtual = "/incs/dbConnection.inc" -->
<!-- #include virtual="/incs/rw.asp" -->

<%
	''On Error Resume Next
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	dim exists
	dim ID, IDarray
	dim EVENT_CD,EVENT_CDarray
	dim MATCH_NUM, MATCH_NUMarray
	dim MATCH_START_TIME, MATCH_START_TIMEarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Time" then
		EVENT_CD		 		= Replace(Request("EVENT_CD"),"'","''")
		MATCH_NUM		 		= Replace(Request("MATCH_NUM"),"'","''")
		MATCH_START_TIME 		= Replace(Request("MATCH_START_TIME"),"'","''")

		'check to see if the location code entered exists'
		checkTime()

		if exists <> "Y" then
			sql = "INSERT INTO TIME_TBL("_
				& "EVENT_CD, "_
				& "MATCH_NUM, "_
				& "MATCH_START_TIME)"_
				& " VALUES("_
				& "'" & EVENT_CD & "', "_
				& "'" & MATCH_NUM & "', "_
				& "'" & MATCH_START_TIME & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Time.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Time.asp")


	ElseIf Request("submitChoice") = "Modify Time(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		MATCH_NUM = FixEmptyCell(Request("MATCH_NUM"))
		MATCH_NUMarray = split(MATCH_NUM,", ")

		MATCH_START_TIME = FixEmptyCell(Request("MATCH_START_TIME"))
		MATCH_START_TIMEarray = split(MATCH_START_TIME,", ")

		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			checkTime()

			if exists <> "Y" then
				sql = "UPDATE TIME_TBL SET "_
					& "EVENT_CD = '" & EVENT_CDarray(i) & "', "_
					& "MATCH_NUM = '" & MATCH_NUMarray(i) & "', "_
					& "MATCH_START_TIME = '" & MATCH_START_TIMEarray(i) & "' "_
					& "WHERE TIME_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & MATCH_NUMarray(i) & "/" & MATCH_START_TIMEarray(i)
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Time.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Time.asp")
	End If


'******************************************'

Sub checkTime()

	If Request("submitChoice") = "Modify Time(s)" then
	''not completed yet
		sql = "SELECT TIME_ID FROM TIME_TBL "_
			& "WHERE EVENT_CD = '" & EVENT_CDarray(i) & "' "_
			& "AND MATCH_NUM = '" & MATCH_NUMarray(i) & "' "_
			& "AND MATCH_START_TIME = '" & MATCH_START_TIMEarray(i) & "' "_
			& "AND TIME_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT TIME_ID FROM TIME_TBL WHERE EVENT_CD = '" & EVENT_CD & "' "_
			& "AND MATCH_NUM = '" & MATCH_NUM & "' "_
			& "AND MATCH_START_TIME = '" & MATCH_START_TIME & "'"
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

Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function
%>