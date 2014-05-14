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
	dim RATER,RATERarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Rater" then
		RATER 		= Replace(Request("RATER"),"'","''")

		'check to see if the location code entered exists'
		checkRater()

		if exists <> "Y" then
			sql = "INSERT INTO RATER_TBL("_
				& "PERSON_ID)"_
				& " VALUES("_
				& "'" & RATER & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Rater.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Rater.asp")


	ElseIf Request("submitChoice") = "Modify Rater(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		RATER = FixEmptyCell(Request("RATER"))
		RATERarray = split(RATER,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			checkRater()

			if exists <> "Y" then
				sql = "UPDATE RATER_TBL SET "_
					& "PERSON_ID = '" & RATERarray(i) & "' "_
					& "WHERE PERSON_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & IDarray(i)
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Rater.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Rater.asp")
	End If


'******************************************'

Sub checkRater()

	If Request("submitChoice") = "Modify Rater(s)" then
		sql = "SELECT PERSON_ID FROM RATER_TBL WHERE PERSON_ID = '" & RATERarray(i) & "' "_
			& "AND PERSON_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT PERSON_ID FROM RATER_TBL WHERE PERSON_ID = '" & RATER & "' "
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