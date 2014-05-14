<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Match Game Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor="#FFFFFF">
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Match Game Administration</u></b>
	</font>
	</div>

	<br />

<%

class matchgame
''private segment, org


sub class_initialize()
	if request.totalbytes=0 then

	else

		if request("btnsave") <> "" then

			dim i

			''deal with deletes
			for each i in request.form

				if i = "REMOVE" then
					arrRemove = split(request("REMOVE"), ", ")

					for j = 0 to UBound(arrRemove)
						arrRemoveParts = split(arrRemove(j),"_")

						SQL = 	"DELETE FROM MATCH_GAME_TBL " & _
								"WHERE MATCH_ID='" & arrRemoveParts(0) & "' AND GAME_NUM='" & arrRemoveParts(1) & "'"
						cn.Execute(SQL)

					next

				end if

			next

			for each i in request.form
				''rw(i & "<br />")
				arrid = split(i, "_")
				''rw(arrid(0) & "<br />")
				''rw(arrid(1) & "<br />")
				''rw(arrid(2) & "<br />")

				if arrid(0) = "GAMENUM" then

					update = "Y"
					''validate numeric data first
					if not (isNumeric(request("GAMENUM_" & arrid(1) & "_" & arrid(2))) and isNumeric(request("TEAM1SCORE_" & arrid(1) & "_" & arrid(2)))  and isNumeric(request("TEAM2SCORE_" & arrid(1) & "_" & arrid(2)))) then
						update = "N"
						Exit For
					end if

					if request("DBL_FORFEIT_IND_" & arrid(1) & "_" & arrid(2)) <> "" then
						TEAM1_SCORE = 0
						TEAM2_SCORE = 0
						DBL_FORFEIT_IND = "Y"
					else
						TEAM1_SCORE = request("TEAM1SCORE_" & arrid(1) & "_" & arrid(2))
						TEAM2_SCORE = request("TEAM2SCORE_" & arrid(1) & "_" & arrid(2))
						DBL_FORFEIT_IND = "N"
					end if

					if request("REFFORFEITIND_" & arrid(1) & "_" & arrid(2)) <> "" then
						REFFORFEIT = "Y"
					else
						REFFORFEIT = "N"
					end if

					if update = "Y" then
						sql = 	"update MATCH_GAME_TBL " & _
								"set [GAME_NUM] = '" & request("GAMENUM_" & arrid(1) & "_" & arrid(2)) & "', " & _
								"[TEAM1_SCORE] = '" & TEAM1_SCORE & "', " & _
								"[TEAM2_SCORE] = '" & TEAM2_SCORE & "', " & _
								"DBL_FORFEIT_IND = '" & DBL_FORFEIT_IND & "', " & _
								"REF_FORFEIT_IND = '" & REFFORFEIT & "' " & _
								"where MATCH_ID = '" & arrid(1) & "' " & _
								"and GAME_NUM = '" & arrid(2) & "'"
						''response.write sql & "<br />"
						cn.Execute(sql)
					end if

				end if

			next 'i

			if update <> "N" then
				session("message") = session("message") & "Match Games have been updated successfully.<br />"
			else
				session("message") = "<font class='cfontError10'>Make sure all fields hold numeric values.</font>"
			end if

			if request("GAMENUMnew") <> "" and request("MATCHIDnew") <> "" then

				update = "Y"
				''validate numeric data first
				if not (isNumeric(request("GAMENUMnew")) and isNumeric(request("TEAM1SCOREnew")) and isNumeric(request("TEAM2SCOREnew"))) then
					update = "N"
				else
					sql = "select * from MATCH_GAME_TBL mg WHERE MATCH_ID = '" & request("MATCHIDnew") & "' AND GAME_NUM = '" & request("GAMENUMnew") & "'"
					set rs = cn.Execute(sql)

					if not rs.EOF then
						update = "DUP"
					end if
				end if

				if update = "Y" then

					sql = "SELECT TEAM1_TEAM_ID, TEAM2_TEAM_ID " & _
							"from MATCH_SCHEDULE_TBL " & _
							"WHERE MATCH_ID='" & request("MATCHIDnew") & "'"
					set rs = cn.Execute(sql)
					TEAM1_ID = rs(0)
					TEAM2_ID = rs(1)

					if request("DBL_FORFEIT_INDnew") <> "" then
						TEAM1_SCORE = 0
						TEAM2_SCORE = 0
						DBL_FORFEIT_IND = "Y"
					else
						TEAM1_SCORE = request("TEAM1SCOREnew")
						TEAM2_SCORE = request("TEAM2SCOREnew")
						DBL_FORFEIT_IND = "N"
					end if

					if request("REF_FORFEIT_INDnew") <> "" then
						REFFORFEIT = "Y"
					else
						REFFORFEIT = "N"
					end if

					sql = 	"insert into MATCH_GAME_TBL" & _
							"(MATCH_ID, " & _
							"GAME_NUM, " & _
							"TEAM1_ID, " & _
							"TEAM2_ID, " & _
							"TEAM1_SCORE, " & _
							"TEAM2_SCORE, " & _
							"DBL_FORFEIT_IND, " & _
							"REF_FORFEIT_IND)" & _
							" values('" & request("MATCHIDnew") & "', " & _
							"'" & request("GAMENUMnew") & "', " & _
							"'" & TEAM1_ID & "', " & _
							"'" & TEAM2_ID & "', " & _
							"'" & TEAM1_SCORE & "', " & _
							"'" & TEAM2_SCORE & "', " & _
							"'" & DBL_FORFEIT_IND & "', " & _
							"'" & REFFORFEIT & "')"

					MATCHID_SELECTED = request("MATCHIDnew")
					''response.write sql & "<br />"
					cn.Execute(sql)
					session("message") = session("message") & "A new match game has been added.<br />"
				elseif update = "N" then
					session("message") = session("message") & "<font class='cfontError10'>Make sure all fields for the new record hold numeric values.</font>"
				elseif update = "DUP" then
					session("message") = session("message") & "<font class='cfontError10'>The new game number entered already exists for the selected match. Please check your information, and try entering the game information again.</font>"
				end if

			end if

		end if ''btnsave

	end if ''request.totalbytes

	WEEK = request("WEEK")

	if WEEK <> "" then
		selectweek WEEK
		printform WEEK,MATCHID_SELECTED
	else
		selectweek ""
	end if
%>
<!--#include virtual="/incs/fragContact.asp"-->

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub selectweek(WEEK)
	sql = "SELECT w.WEEK_ID, w.WEEK_NUM, w.[DATE], e.EVENT_SHORT_DESC "_
		& "FROM WEEK_TBL w LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
		& "WHERE e.ACTIVE_EVENT_IND IN ('Y','A') "_
		& "ORDER BY UPPER(EVENT_SHORT_DESC), w.[DATE]"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsWeekData = rs.GetRows
		rsWeekRows = UBound(rsWeekData,2)
	else
		rw("Error:Missing Weeks.")
		Response.End
	end if
%>
	<table align='center' bgcolor="#FFFFFF">
	<tr><td>
		<font class='cfont8'><b>
		<ul>
		<li>Please select an event/week from the first drop down list. Then select a match from the drop down.<br />
			Enter the game number and the score, making sure to match team 1 and team 2 according to the drop down value.<br />
			Then click 'Submit' to enter the new game information.
		</li>
		<li>To update a game's score, re-enter the scores/game number as needed and click 'Submit'.</li>
		<li>To remove a game's score, click the appropriate checkbox and click 'Submit'.</li>
		<li>Both double forfeit scores should be entered as 0 with the checkbox selected.</li>
		<li>Check the ref penalty box if the ref team did not ref the match.</li>
		</ul>
		</b></font>
	</td></tr>
	</table>

	<br />
	<form name='weekChoice'>
	<div align='center'>
		<font class='cfont10'>
		Event/Week:
		<select name='WEEK' onChange="changePage();">
		<option value=''>-select-</option>
		<%
			For i = 0 to rsWeekRows
				rw("<option value='" & rsWeekData(0,i) & "'")
					If CStr(WEEK) = CStr(rsWeekData(0,i)) then
					rw(" selected")
				End If
					rw(">" & rsWeekData(3,i) & " - Week " & rsWeekData(1,i) & " - " & rsWeekData(2,i) & "</option>")
			Next
		%>
		</select>
		</font>
	</div>
	</form>

	<script language='Javascript'>
	<!--

		function changePage()
		{
			choice = weekChoice.WEEK.options[weekChoice.WEEK.selectedIndex].value;
			//alert(choice);
			if(choice != "")
			{
				weekChoice.submit();
			}
		}
	//-->
	</script>

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub printform(WEEK,MATCHID_SELECTED)
	sql = "MATCH_GAME @WEEK='" & WEEK & "'"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsMGData = rs.GetRows
		rsMGRows = UBound(rsMGData,2)
	else
		rsMGRows = -1
	end if

''	rw(sql)
''	Response.End

	sql = "MATCH_SCHEDULE_MATCHLIST @WEEK='" & WEEK & "'"
	set rs = cn.Execute(sql)

	if session("message") <> "" then
		response.write "<div class=""cfontSuccess10"" align='center'>" & session("message") & "</div>"
		session("message") = ""
	end if

%>

<form method="post" name='MODIFY'>

<table width='500' bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>
<tr bgcolor='#000066'>
<th><font class='cfontWhite10'>Remove?</font></th>
<th><font class='cfontWhite10'>Time</font></th>
<th><font class='cfontWhite10'>Game #</font></th>
<th><font class='cfontWhite10'>Team 1<br />(Div/Code/Name)</font></th>
<th><font class='cfontWhite10'>Team 2<br />(Div/Code/Name)</font></th>
<th><font class='cfontWhite10'>Team 1 Score</font></th>
<th><font class='cfontWhite10'>Team 2 Score</font></th>
<th><font class='cfontWhite10'>Double Forfeit?</font></th>
<th><font class='cfontWhite10'>Ref Penalty?</font></th>
</tr>

<tr>
<th colspan="9"><font class='cfont10'>New Game</font></th>
</tr>

<tr bgcolor='#FFFFFF'>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td><input type="text" maxlength="1" style="width:30px;" name="GAMENUMnew" id="GAMENUMnew" /></td>
<td colspan='2'><%makedd rs,"MATCHIDnew",MATCHID_SELECTED %></td>
<td><input type="text" maxlength="2" style="width:30px;" name="TEAM1SCOREnew" id="TEAM1SCOREnew" /></td>
<td><input type="text" maxlength="2" style="width:30px;" name="TEAM2SCOREnew" id="TEAM2SCOREnew" /></td>
<td align='center'><input type="checkbox" name="DBL_FORFEIT_INDnew" id="DBL_FORFEIT_INDnew" value='Y' /></td>
<td align='center'><input type="checkbox" name="REF_FORFEIT_INDnew" id="REF_FORFEIT_INDnew" value='Y' /></td>
</tr>

<tr>
<td colspan="9"><input type="submit" name="btnsave" value="Submit" class="btn" /></td>
</tr>

<%
	For i = 0 to rsMGRows

		ODD_ROW = NOT ODD_ROW

		If ODD_ROW Then
			BGCOLOR = "#FFFFFF"
		Else
			BGCOLOR = "#F0F0F0"
		End If
%>
		<tr bgcolor='<%=BGCOLOR%>'>
		<td align='center'><input type='checkbox' name='REMOVE' value='<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>' /></td>
		<td><font class='cfont10'><%=rsMGData(8,i)%></font></td>
		<td><input type="text" maxlength='1' style="width:30px;" name="GAMENUM_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" id="GAMENUM_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" value="<%=rsMGData(1,i)%>" /></td>
		<td><font class='cfont10'><%=rsMGData(6,i)%></font></td>
		<td><font class='cfont10'><%=rsMGData(7,i)%></font></td>
		<td><input type="text" maxlength='2' style="width:30px;" name="TEAM1SCORE_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" id="TEAM1SCORE_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" value="<%=rsMGData(4,i)%>" /></td>
		<td><input type="text" maxlength='2' style="width:30px;" name="TEAM2SCORE_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" id="TEAM2SCORE_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" value="<%=rsMGData(5,i)%>" /></td>
		<td align='center'><input type="checkbox" name="DBL_FORFEIT_IND_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" id="DBL_FORFEIT_IND_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" value='Y' <%If rsMGData(9,i)  = "Y" then rw(" checked") End If%> /></td>
		<td align='center'><input type="checkbox" name="REFFORFEITIND_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" id="REFFORFEITIND_<%=rsMGData(0,i)%>_<%=rsMGData(1,i)%>" value='Y' <%If rsMGData(10,i)  = "Y" then rw(" checked") End If%> /></td>
		</tr>
<%
	Next
%>

</table>

</form>

<%

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub makedd(rs, name, selected)
%>
<select name="<%=name%>" id="<%=name%>">
<option value="">- Select A Match -</option>
<%
	while not rs.eof

%>

		<option<%If Cstr(selected) = CStr(rs(0)) Then rw(" selected") End If%> value="<%=rs(0)%>">(<%=rs(1)%>) <%=rs(2)%> vs. <%=rs(3)%></option>

<%
		rs.movenext
	wend

%>
</select>

<%
end sub

end class


set ag = new matchgame

%>


