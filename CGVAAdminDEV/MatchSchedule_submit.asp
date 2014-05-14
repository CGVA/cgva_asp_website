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
	dim WEEK_ID, WEEK_IDarray
	dim TIME_ID, TIME_IDarray
	dim COURT_NUM, COURT_NUMarray
	dim TEAM1_TEAM_ID, TEAM1_TEAM_IDarray
	dim TEAM2_TEAM_ID, TEAM2_TEAM_IDarray
	dim REF_TEAM_ID, REF_TEAM_IDarray
	dim TEAM1_MVP_ID, TEAM1_MVP_IDarray
	dim TEAM2_MVP_ID, TEAM2_MVP_IDarray
	dim TEAM1_INCLUDE_DIV_STATS_IND, TEAM1_INCLUDE_DIV_STATS_INDarray
	dim TEAM2_INCLUDE_DIV_STATS_IND, TEAM2_INCLUDE_DIV_STATS_INDarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Match Schedule" then
		WEEK_ID  						= Replace(Request("WEEK_ID"),"'","''")
		Session("WEEK_ID")				= Replace(Request("WEEK_ID"),"'","''")
		TIME_ID      					= Replace(Request("TIME_ID"),"'","''")
		Session("TIME_ID")     			= Replace(Request("TIME_ID"),"'","''")
		COURT_NUM      					= Replace(Request("COURT_NUM"),"'","''")
		Session("COURT_NUM")      		= Replace(Request("COURT_NUM"),"'","''")
		TEAM1_TEAM_ID  					= Replace(Request("TEAM1_TEAM_ID"),"'","''")
		Session("TEAM1_TEAM_ID")  		= Replace(Request("TEAM1_TEAM_ID"),"'","''")
		TEAM2_TEAM_ID   				= Replace(Request("TEAM2_TEAM_ID"),"'","''")
		Session("TEAM2_TEAM_ID")  		= Replace(Request("TEAM2_TEAM_ID"),"'","''")
		REF_TEAM_ID						= Replace(Request("REF_TEAM_ID"),"'","''")
		Session("REF_TEAM_ID")			= Replace(Request("REF_TEAM_ID"),"'","''")
		''TEAM1_MVP_ID
		''TEAM2_MVP_ID
		TEAM1_INCLUDE_DIV_STATS_IND  	= Replace(Request("TEAM1_INCLUDE_DIV_STATS_IND"),"'","''")
		TEAM2_INCLUDE_DIV_STATS_IND		= Replace(Request("TEAM2_INCLUDE_DIV_STATS_IND"),"'","''")

		'make sure team1,team2 and ref team are unique'
		''already done with front end javascript
		''if 	TEAM1_TEAM_ID = TEAM2_TEAM_ID or TEAM1_TEAM_ID = REF_TEAM_ID or TEAM2_TEAM_ID = REF_TEAM_ID then
		''	Session("admin") = "insertDup"
		''	Response.Redirect("MatchSchedule.asp")
		''end if

		'check to see if this match entered already exists'
		checkMatchSchedule()

		''verify that this matchup is for teams in the same event
		sql = "SELECT COUNT(DISTINCT(b.EVENT_CD)) FROM TEAM_TBL a "_
			& "LEFT JOIN DIVISION_TBL b ON a.DIVISION_ID = b.DIVISION_ID "_
			& "LEFT JOIN TIME_TBL t ON b.EVENT_CD = t.EVENT_CD "_
			& "WHERE a.TEAM_ID IN ('" & TEAM1_TEAM_ID & "','" & TEAM2_TEAM_ID & "','" & REF_TEAM_ID & "') "_
			& "OR t.TIME_ID ='" & TIME_ID & "'"
		rw(sql & "<br />")
		set rs = cn.Execute(sql)

		if rs(0) <> 1 then
			call closeRSCNConnection()
			Session("admin") = "insertFailEvent"
			Response.Redirect("MatchSchedule.asp")
		end if

		''verify the selected teams arent already scheduled for a match at the selected time
		sql = "SELECT COUNT(*) FROM MATCH_SCHEDULE_TBL  "_
			& "WHERE (TEAM1_TEAM_ID IN ('" & TEAM1_TEAM_ID & "','" & TEAM2_TEAM_ID & "','" & REF_TEAM_ID & "') "_
			& "		or TEAM2_TEAM_ID IN ('" & TEAM1_TEAM_ID & "','" & TEAM2_TEAM_ID & "','" & REF_TEAM_ID & "') "_
			& "		or REF_TEAM_ID IN ('" & TEAM1_TEAM_ID & "','" & TEAM2_TEAM_ID & "','" & REF_TEAM_ID & "')) "_
			& "AND WEEK_ID = '" & WEEK_ID & "' "_
			& "AND TIME_ID = '" & TIME_ID & "'"
		rw(sql & "<br />")
		''Response.End
		set rs = cn.Execute(sql)

		if rs(0) <> 0 then
			call closeRSCNConnection()
			Session("admin") = "insertDup"
			Response.Redirect("MatchSchedule.asp")
		end if

		''verify no other match is scheduled for the selected week/time/court
		sql = "SELECT COUNT(*) FROM MATCH_SCHEDULE_TBL  "_
			& "WHERE WEEK_ID = '" & WEEK_ID & "' "_
			& "AND TIME_ID = '" & TIME_ID & "' "_
			& "AND COURT_NUM = '" & COURT_NUM & "'"
		rw(sql & "<br />")
		''Response.End
		set rs = cn.Execute(sql)

		if rs(0) <> 0 then
			call closeRSCNConnection()
			Session("admin") = "insertDupCourt"
			Response.Redirect("MatchSchedule.asp")
		end if



		if exists <> "Y" then
			sql = "INSERT INTO MATCH_SCHEDULE_TBL("_
				& "WEEK_ID, "_
				& "TIME_ID, "_
				& "COURT_NUM, "_
				& "TEAM1_TEAM_ID, "_
				& "TEAM2_TEAM_ID, "_
				& "REF_TEAM_ID, "_
				& "TEAM1_INCLUDE_DIV_STATS_IND, "_
				& "TEAM2_INCLUDE_DIV_STATS_IND)"_
				& " VALUES("_
				& "'" & WEEK_ID & "', "_
				& "'" & TIME_ID & "', "_
				& "'" & COURT_NUM & "', "_
				& "'" & TEAM1_TEAM_ID & "', "_
				& "'" & TEAM2_TEAM_ID & "', "_
				& "'" & REF_TEAM_ID & "', "_
				& "'" & TEAM1_INCLUDE_DIV_STATS_IND & "', "_
				& "'" & TEAM2_INCLUDE_DIV_STATS_IND & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("MatchSchedule.asp")
		end if

		call closeCNConnection()
		Session("admin") = "insert"
		Response.Redirect("MatchSchedule.asp")


	ElseIf Request("submitChoice") = "Modify Match Schedule(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		WEEK_ID = FixEmptyCell(Request("WEEK_ID"))
		WEEK_IDarray = split(WEEK_ID,", ")

		TIME_ID = FixEmptyCell(Request("TIME_ID"))
		TIME_IDarray = split(TIME_ID,", ")

		COURT_NUM = FixEmptyCell(Request("COURT_NUM"))
		COURT_NUMarray = split(COURT_NUM,", ")

		TEAM1_TEAM_ID = FixEmptyCell(Request("TEAM1_TEAM_ID"))
		TEAM1_TEAM_IDarray = split(TEAM1_TEAM_ID,", ")

		TEAM2_TEAM_ID = FixEmptyCell(Request("TEAM2_TEAM_ID"))
		TEAM2_TEAM_IDarray = split(TEAM2_TEAM_ID,", ")

		REF_TEAM_ID = FixEmptyCell(Request("REF_TEAM_ID"))
		REF_TEAM_IDarray = split(REF_TEAM_ID,", ")

		TEAM1_MVP_ID = FixEmptyCell(Request("TEAM1_MVP_ID"))
		TEAM1_MVP_IDarray = split(TEAM1_MVP_ID,", ")

		TEAM2_MVP_ID = FixEmptyCell(Request("TEAM2_MVP_ID"))
		TEAM2_MVP_IDarray = split(TEAM2_MVP_ID,", ")

		TEAM1_INCLUDE_DIV_STATS_IND = FixEmptyCell(Request("TEAM1_INCLUDE_DIV_STATS_IND"))
		TEAM1_INCLUDE_DIV_STATS_INDarray = split(TEAM1_INCLUDE_DIV_STATS_IND,", ")

		TEAM2_INCLUDE_DIV_STATS_IND = FixEmptyCell(Request("TEAM2_INCLUDE_DIV_STATS_IND"))
		TEAM2_INCLUDE_DIV_STATS_INDarray = split(TEAM2_INCLUDE_DIV_STATS_IND,", ")


		''validate MVP belongs to the selected team (if entered)
		For i = 0 to UBound(IDarray)
			if StrComp(CStr(TEAM1_MVP_IDarray(i)),"0",1) <> 0 then
				SQL = "SELECT TEAM_ID FROM TEAM_MEMBER_TBL WHERE PERSON_ID = '" & TEAM1_MVP_IDarray(i) & "' AND TEAM_ID = '" & TEAM1_TEAM_IDarray(i) & "'"
				set rs = cn.Execute(SQL)

				''member not on selected team, so reset MVP
				if rs.EOF then
					TEAM1_MVP_IDarray(i) = 0
					notUpdated = "An MVP was selected for a team for which s/he isn't a member (this selection is not saved).<br />"
				end if
			end if

			if StrComp(CStr(TEAM2_MVP_IDarray(i)),"0",1) <> 0 then
				SQL = "SELECT TEAM_ID FROM TEAM_MEMBER_TBL WHERE PERSON_ID = '" & TEAM2_MVP_IDarray(i) & "' AND TEAM_ID = '" & TEAM2_TEAM_IDarray(i) & "'"
				set rs = cn.Execute(SQL)

				''member not on selected team, so reset MVP
				if rs.EOF then
					TEAM2_MVP_IDarray(i) = 0
					notUpdated = notUpdated  & "An MVP was selected for a team for which s/he isn't a member (this selection is not saved).<br />"
				end if
			end if

		Next

		For i = 0 to UBound(IDarray)

			checkMatchSchedule()

			if exists <> "Y" then

				if StrComp(CStr(TEAM1_MVP_IDarray(i)),"0",1) = 0 then
					TEAM1_MVP_IDarray(i) = "NULL"
				end if

				if StrComp(CStr(TEAM2_MVP_IDarray(i)),"0",1) = 0 then
					TEAM2_MVP_IDarray(i) = "NULL"
				end if

				sql = "UPDATE MATCH_SCHEDULE_TBL SET "_
					& "WEEK_ID = '" & WEEK_IDarray(i) & "', "_
					& "TIME_ID = '" & TIME_IDarray(i) & "', "_
					& "COURT_NUM = '" & COURT_NUMarray(i) & "', "_
					& "TEAM1_TEAM_ID = '" & TEAM1_TEAM_IDarray(i) & "', "_
					& "TEAM2_TEAM_ID = '" & TEAM2_TEAM_IDarray(i) & "', "_
					& "REF_TEAM_ID = '" & REF_TEAM_IDarray(i) & "', "_
					& "TEAM1_MVP_ID = " & TEAM1_MVP_IDarray(i) & ", "_
					& "TEAM2_MVP_ID = " & TEAM2_MVP_IDarray(i) & ", "_
					& "TEAM1_INCLUDE_DIV_STATS_IND = '" & TEAM1_INCLUDE_DIV_STATS_INDarray(i) & "', "_
					& "TEAM2_INCLUDE_DIV_STATS_IND = '" & TEAM2_INCLUDE_DIV_STATS_INDarray(i) & "' "_
					& "WHERE MATCH_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & "(court/time ID): " & COURT_NUMarray(i) & "/" & TIME_IDarray(i) & " was not updated.<br />"
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("MatchSchedule.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("MatchSchedule.asp")
	End If


'******************************************'

Sub checkMatchSchedule()

	'see if this match has already been entered
	If Request("submitChoice") = "Modify Match Schedule(s)" then
		sql = "SELECT MATCH_ID FROM MATCH_SCHEDULE_TBL "_
			& "WHERE WEEK_ID = '" & WEEK_IDarray(i) & "' "_
			& "AND TIME_ID = '" & TIME_IDarray(i) & "' "_
			& "AND ((TEAM1_TEAM_ID = '" & TEAM1_TEAM_IDarray(i) & "' AND TEAM2_TEAM_ID = '" & TEAM2_TEAM_IDarray(i) & "') "_
			& "OR (TEAM2_TEAM_ID = '" & TEAM1_TEAM_IDarray(i) & "' AND TEAM1_TEAM_ID = '" & TEAM2_TEAM_IDarray(i) & "')) "_
			& "AND REF_TEAM_ID = '" & REF_TEAM_IDarray(i) & "' "_
			& "AND MATCH_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT MATCH_ID FROM MATCH_SCHEDULE_TBL "_
			& "WHERE WEEK_ID = '" & WEEK_ID & "' "_
			& "AND TIME_ID = '" & TIME_ID & "' "_
			& "AND ((TEAM1_TEAM_ID = '" & TEAM1_TEAM_ID & "' AND TEAM2_TEAM_ID = '" & TEAM2_TEAM_ID & "') "_
			& "OR (TEAM2_TEAM_ID = '" & TEAM1_TEAM_ID & "' AND TEAM1_TEAM_ID = '" & TEAM2_TEAM_ID & "')) "_
			& "AND REF_TEAM_ID = '" & REF_TEAM_ID & "'"
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