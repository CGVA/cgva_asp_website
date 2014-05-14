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
	dim DIVISION_ID, DIVISION_IDarray
	dim TEAM_CD, TEAM_CDarray
	dim TEAM_NAME, TEAM_NAMEarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Team" then
		DIVISION_ID		 	= Replace(Request("DIVISION_ID"),"'","''")
		TEAM_CD		 		= Replace(Request("TEAM_CD"),"'","''")
		TEAM_NAME 			= Replace(Request("TEAM_NAME"),"'","''")

		'check to see if the team id entered exists'
		checkTeam()

		if exists <> "Y" then
			sql = "INSERT INTO TEAM_TBL("_
				& "DIVISION_ID, "_
				& "TEAM_CD, "_
				& "TEAM_NAME)"_
				& " VALUES("_
				& "'" & DIVISION_ID & "', "_
				& "'" & TEAM_CD & "', "_
				& "'" & TEAM_NAME & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Team.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Team.asp")


	ElseIf Request("submitChoice") = "Modify Team(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		DIVISION_ID = FixEmptyCell(Request("DIVISION_ID"))
		DIVISION_IDarray = split(DIVISION_ID,", ")

		TEAM_CD = FixEmptyCell(Replace(Request("TEAM_CD"),"'","''"))
		TEAM_CDarray = split(TEAM_CD,", ")

		TEAM_NAME = FixEmptyCell(Replace(Request("TEAM_NAME"),"'","''"))
		TEAM_NAMEarray = split(TEAM_NAME,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			checkTeam()

			if exists <> "Y" then
				sql = "UPDATE TEAM_TBL SET "_
					& "DIVISION_ID = '" & DIVISION_IDarray(i) & "', "_
					& "TEAM_CD = '" & TEAM_CDarray(i) & "', "_
					& "TEAM_NAME = '" & TEAM_NAMEarray(i) & "' "_
					& "WHERE TEAM_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & TEAM_CDarray(i) & ":" & TEAM_NAMEarray(i) & "<br />"
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Team.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Team.asp")
	End If


'******************************************'

Sub checkTeam()

	If Request("submitChoice") = "Modify Team(s)" then
	''not completed yet
		sql = "SELECT UPPER(TEAM_NAME) FROM TEAM_TBL WHERE DIVISION_ID = '" & UCase(DIVISION_IDarray(i)) & "' "_
			& "AND UPPER(TEAM_NAME) = '" & UCase(TEAM_NAMEarray(i)) & "' "_
			& "AND TEAM_ID <> '" & UCase(IDarray(i)) & "' "_
			& "UNION "_
			& "SELECT UPPER(TEAM_CD) FROM TEAM_TBL WHERE DIVISION_ID = '" & UCase(DIVISION_IDarray(i)) & "' "_
			& "AND UPPER(TEAM_CD) = '" & UCase(TEAM_CDarray(i)) & "' "_
			& "AND TEAM_ID <> '" & UCase(IDarray(i)) & "'"

	Else
		sql = "SELECT UPPER(TEAM_NAME) FROM TEAM_TBL WHERE DIVISION_ID = '" & DIVISION_ID & "' "_
			& "AND UPPER(TEAM_NAME) = '" & UCase(TEAM_NAME) & "' "_
			& "UNION "_
			& "SELECT UPPER(TEAM_CD) FROM TEAM_TBL WHERE DIVISION_ID = '" & DIVISION_ID & "' "_
			& "AND UPPER(TEAM_CD) = '" & UCase(TEAM_CD) & "' "_

	End If

	rw(sql & "<br />")
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