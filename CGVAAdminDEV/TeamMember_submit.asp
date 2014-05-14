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
	dim PERSON_ID, PERSON_IDarray
	dim CAPTAIN_IND, CAPTAIN_INDarray
''	dim CERTIFIED_REF_IND, CERTIFIED_REF_INDarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If


	''rw("here")
	''Response.End

	EDT = Request("EDT")
	EDTArray = Split(EDT,"_")
	EVENT_CD = EDTArray(0)
	TEAM = EDTArray(1)

	'determine which record to modify based on the ID number of the record'
	PERSON_ID = Request("PERSON_ID")
	PERSON_IDarray = split(PERSON_ID,", ")

	CAPTAIN_IND = FixEmptyCell(Request("CAPTAIN_IND"))
	CAPTAIN_INDarray = split(CAPTAIN_IND,", ")

''	CERTIFIED_REF_IND = FixEmptyCell(Replace(Request("CERTIFIED_REF_IND"),"'","''"))
''	CERTIFIED_REF_INDarray = split(CERTIFIED_REF_IND,", ")

	exists = ""
	'check to see if one of the login ID entered exists on another team'
	For i = 0 to UBound(PERSON_IDarray)

		rw(i & ": " & PERSON_IDarray(i) & "<br />")

		if PERSON_IDarray(i) <> "" then

			sql = "SELECT a.PERSON_ID "_
				& "FROM TEAM_MEMBER_TBL a "_
				& "LEFT JOIN TEAM_TBL t ON a.TEAM_ID = t.TEAM_ID "_
				& "LEFT JOIN DIVISION_TBL d ON t.DIVISION_ID = d.DIVISION_ID "_
				& "WHERE a.PERSON_ID = '" & PERSON_IDarray(i) & "' "_
				& "AND t.TEAM_ID <> '" & TEAM & "' "_
				& "AND d.EVENT_CD = '" & EVENT_CD & "'"

				''& "LEFT JOIN EVENT_TBL e ON d.EVENT_CD = e.EVENT_CD "_

			''rw(sql & "<br />")
			''Response.End

			set rs = cn.Execute(sql)

			If not rs.EOF then
				exists = exists & PERSON_IDarray(i) & ", "
			End If

		end if

	Next

	if exists = "" then

		''clear current team
		sql = "DELETE FROM TEAM_MEMBER_TBL WHERE TEAM_ID='" & TEAM & "'"
		rw(sql & "<br />")
		cn.Execute(sql)

		For i = 0 to UBound(PERSON_IDarray)

			if PERSON_IDarray(i) <> "" then

				''add all new members, verifying the person wasnt selected twice on the entry screen
				sql	= "SELECT * FROM TEAM_MEMBER_TBL "_
					& "WHERE TEAM_ID='" & TEAM & "' "_
					& "AND PERSON_ID='" & PERSON_IDarray(i) & "'"
				set rs = cn.Execute(sql)

				if rs.EOF then

'					sql = "INSERT INTO TEAM_MEMBER_TBL(TEAM_ID,PERSON_ID,CAPTAIN_IND,CERTIFIED_REF_IND) VALUES("_
'						& "'" & TEAM & "', "_
'						& "'" & PERSON_IDarray(i) & "', "_
'						& "'" & CAPTAIN_INDarray(i) & "', "_
'						& "'" & CERTIFIED_REF_INDarray(i) & "')"

					sql = "INSERT INTO TEAM_MEMBER_TBL(TEAM_ID,PERSON_ID,CAPTAIN_IND) VALUES("_
						& "'" & TEAM & "', "_
						& "'" & PERSON_IDarray(i) & "', "_
						& "'" & CAPTAIN_INDarray(i) & "')"
					''rw(sql & "<br />")
					''Response.End

					cn.Execute(sql)

				end if

			end if

		Next

	else
		notUpdated = Left(exists,Len(exists)-2)
	end if

	call closeCNConnection()
	Session("admin") = "insert"
	Response.Redirect("TeamMember.asp?notupdated=" & notUpdated)



'******************************************'
Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function
'******************************************'

%>