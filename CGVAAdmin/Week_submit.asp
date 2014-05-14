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
    dim WEEK_NUM, WEEK_NUMarray
	dim vDATE, vDATEarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Week" then
		EVENT_CD		 	= Replace(Request("EVENT_CD"),"'","''")
		WEEK_NUM		 	= Replace(Request("WEEK_NUM"),"'","''")
		vDATE		 		= Replace(Request("DATE"),"'","''")

		'check to see if the location code entered exists'
		checkWeek()

		if exists <> "Y" then
			sql = "INSERT INTO WEEK_TBL("_
				& "EVENT_CD, "_
				& "WEEK_NUM, "_
				& "[DATE])"_
				& " VALUES("_
				& "'" & EVENT_CD & "', "_
				& "'" & WEEK_NUM & "', "_
				& "'" & vDATE & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Week.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Week.asp")


	ElseIf Request("submitChoice") = "Modify Week(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		WEEK_NUM = FixEmptyCell(Request("WEEK_NUM"))
		WEEK_NUMarray = split(WEEK_NUM,", ")

		vDATE = FixEmptyCell(Request("DATE"))
		vDATEarray = split(vDATE,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)

			checkWeek()

			if exists <> "Y" then
				sql = "UPDATE WEEK_TBL SET "_
					& "EVENT_CD = '" & EVENT_CDarray(i) & "', "_
					& "WEEK_NUM = '" & WEEK_NUMarray(i) & "', "_
					& "[DATE] = '" & vDATEarray(i) & "' "_
					& "WHERE WEEK_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & WEEK_NUMarray(i) & "/" & vDATEarray(i)
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Week.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Week.asp")
	End If


'******************************************'

Sub checkWeek()

	If Request("submitChoice") = "Modify Week(s)" then
	''not completed yet
		sql = "SELECT WEEK_ID FROM WEEK_TBL "_
			& "WHERE EVENT_CD = '" & EVENT_CDarray(i) & "' "_
			& "AND WEEK_NUM = '" & WEEK_NUMarray(i) & "' "_
			& "AND [DATE] = '" & vDATEarray(i) & "' "_
			& "AND WEEK_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT WEEK_ID FROM WEEK_TBL WHERE EVENT_CD = '" & EVENT_CD & "' "_
			& "AND WEEK_NUM = '" & WEEK_NUM & "' "_
			& "AND [DATE] = '" & vDATE & "'"
	End If

	rw(sql)
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