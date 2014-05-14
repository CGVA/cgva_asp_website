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

	sql = "SELECT w.WEEK_ID, w.WEEK_NUM, w.[DATE], e.EVENT_SHORT_DESC "_
		& "FROM WEEK_TBL w LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
		& "WHERE e.ACTIVE_EVENT_IND IN ('A','Y') "_
		& "ORDER BY UPPER(EVENT_SHORT_DESC), w.[DATE]"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsWeekData = rs.GetRows
		rsWeekRows = UBound(rsWeekData,2)
	else
		rw("Error:Missing Weeks.<p />" & sql)
		Response.End
	end if

	sql = "SELECT a.EVENT_SHORT_DESC, b.MATCH_START_TIME, b.TIME_ID "_
		& "FROM TIME_TBL b LEFT JOIN EVENT_TBL a ON b.EVENT_CD = a.EVENT_CD "_
		& "WHERE a.ACTIVE_EVENT_IND IN ('A','Y') "_
		& "ORDER BY UPPER(EVENT_SHORT_DESC), b.MATCH_START_TIME"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsTimeData = rs.GetRows
		rsTimeRows = UBound(rsTimeData,2)
	else
		rw("Error:Missing Times.<p />" & sql)
		Response.End
	end if

''	sql = "SELECT a.TEAM_ID, c.EVENT_SHORT_DESC + '/' + b.DIVISION_CD + '/' + a.TEAM_CD + '(' + a.TEAM_NAME + ')' as 'INFORMATION' "_
	sql = "SELECT a.TEAM_ID, c.EVENT_SHORT_DESC + '/' + b.DIVISION_CD + '/' + a.TEAM_CD as 'INFORMATION' "_
		& "FROM TEAM_TBL a LEFT JOIN DIVISION_TBL b ON a.DIVISION_ID = b.DIVISION_ID "_
		& "LEFT JOIN EVENT_TBL c ON b.EVENT_CD = c.EVENT_CD "_
		& "WHERE c.ACTIVE_EVENT_IND IN ('A','Y') "_
		& "ORDER BY INFORMATION"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsTeamData = rs.GetRows
		rsTeamRows = UBound(rsTeamData,2)
	else
		rw("Error:Missing Teams.<p />" & sql)
		Response.End
	end if


%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Match Schedule Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{


	if(INSERT.COURT_NUM.value == "" || !isInteger(INSERT.COURT_NUM.value))
	{
		alert("Please enter a numeric court number.");
		INSERT.COURT_NUM.focus();
		return false;
	}

	if(INSERT.TEAM1_TEAM_ID.selectedIndex == INSERT.TEAM2_TEAM_ID.selectedIndex || INSERT.TEAM1_TEAM_ID.selectedIndex == INSERT.REF_TEAM_ID.selectedIndex || INSERT.TEAM2_TEAM_ID.selectedIndex == INSERT.REF_TEAM_ID.selectedIndex)
	{
		alert("Team 1/Team 2/Ref Team must be unique.");
		INSERT.TEAM1_TEAM_ID.focus();
		return false;
	}

	//alert("error check completed.");
	//return false;
	return true;
}


////////////////////////////////////////////////////////////////////////////

function validateModify()
{
	var selectedValue = "x";
	if(MODIFY.ID.length > 1)
	{

		for(i = 0;i < MODIFY.ID.length;i++)
		{

			if(MODIFY.ID[i].checked)
			{
				selectedValue = i;

				if(MODIFY.DIVISION_CD[selectedValue].value == "")
				{
					alert("Please enter a division code for the record being modified.");
					MODIFY.DIVISION_CD[selectedValue].focus();
					return false;
				}

				if(MODIFY.DIVISION_DESC[selectedValue].value == "")
				{
					alert("Please enter a division description for the record being modified.");
					MODIFY.DIVISION_DESC[selectedValue].focus();
					return false;
				}
			}

		}

		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID[0].focus();
			return false;
		}

	}
	else
	{

		if(MODIFY.ID.checked)
		{
			selectedValue = 1;
		}


		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID.focus();
			return false;
		}


		if(MODIFY.DIVISION_CD.value == "")
		{
			alert("Please enter a division code for the record being modified.");
			MODIFY.DIVISION_CD.focus();
			return false;
		}

		if(MODIFY.DIVISION_DESC.value == "")
		{
			alert("Please enter a division description for the record being modified.");
			MODIFY.DIVISION_DESC.focus();
			return false;
		}


	}

	//alert("error check completed.");
	//return false;
	return true;
}

function extracheck(obj)
{
	If(obj.disabled == true)
	{
		return !obj.disabled;
	}
	Else
	{
		return obj.disabled;
	}
}

function disableCheck(obj)
{

	rowValue = obj.value;
	var idValue = eval("MODIFY.ID" + rowValue);
	var WEEK_ID = eval("MODIFY.WEEK_ID" + rowValue);
	var TIME_ID = eval("MODIFY.TIME_ID" + rowValue);
	var COURT_NUM = eval("MODIFY.COURT_NUM" + rowValue);
	var TEAM1_TEAM_ID = eval("MODIFY.TEAM1_TEAM_ID" + rowValue);
	var TEAM2_TEAM_ID = eval("MODIFY.TEAM2_TEAM_ID" + rowValue);
	var REF_TEAM_ID = eval("MODIFY.REF_TEAM_ID" + rowValue);
	var TEAM1_MVP_ID = eval("MODIFY.TEAM1_MVP_ID" + rowValue);
	var TEAM2_MVP_ID = eval("MODIFY.TEAM2_MVP_ID" + rowValue);
	var TEAM1_INCLUDE_DIV_STATS_IND = eval("MODIFY.TEAM1_INCLUDE_DIV_STATS_IND" + rowValue);
	var TEAM2_INCLUDE_DIV_STATS_IND = eval("MODIFY.TEAM2_INCLUDE_DIV_STATS_IND" + rowValue);

	if(idValue.checked)
	{
		WEEK_ID.disabled = false;
		TIME_ID.disabled = false;
		COURT_NUM.disabled = false;
		TEAM1_TEAM_ID.disabled = false;
		TEAM2_TEAM_ID.disabled = false;
		REF_TEAM_ID.disabled = false;
		TEAM1_MVP_ID.disabled = false;
		TEAM2_MVP_ID.disabled = false;
		TEAM1_INCLUDE_DIV_STATS_IND.disabled = false;
		TEAM2_INCLUDE_DIV_STATS_IND.disabled = false;
	}else{
		WEEK_ID.disabled = true;
		TIME_ID.disabled = true;
		COURT_NUM.disabled = true;
		TEAM1_TEAM_ID.disabled = true;
		TEAM2_TEAM_ID.disabled = true;
		REF_TEAM_ID.disabled = true;
		TEAM1_MVP_ID.disabled = true;
		TEAM2_MVP_ID.disabled = true;
		TEAM1_INCLUDE_DIV_STATS_IND.disabled = true;
		TEAM2_INCLUDE_DIV_STATS_IND.disabled = true;
	}

}

function checkAll(field)
{
	for (i = 0; i < field.length; i++)
	{
		field[i].checked = true ;
		disableCheck(field[i]);
	}
}

function uncheckAll(field)
{
	for (i = 0; i < field.length; i++)
	{
		field[i].checked = false ;
		disableCheck(field[i]);
	}
}


//-->
</script>

<!-- #include virtual="/incs/integerCheck.js" -->
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Match Schedule Administration</u></b>
	</font>
	</div>

	<br />

	<%
		'call subroutines'
		If Request("choice") = "" then
			Call chooseOption()
			Call closePage()
		ElseIf Request("choice") = "modify" then
			Call modifyRecord()
			Call closePage()
		ElseIf Request("choice") = "insert" then
			Call insertRecord()
			Call closePage()
		End If


	'******************************************'

	Sub chooseOption()
	%>

		<form name="CHOICE" method="POST" action="MatchSchedule.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Match Schedule
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Match Schedule
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<input type="submit" value="Next ->">
		</td>
		</tr>

		</table>

		</form>

		<table align='center' border='0' cellpadding='3'>


		<%
			If Session("admin") = "insert" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontSuccess10'>")
					rw("<b>The new Match Schedule was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A Match Schedule with the same week/time/teams already exists.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFailEvent" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>The selected teams/time are not in the same event.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertDup" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>At least one of the selected teams is already scheduled for the selected week/time.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertDupCourt" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A match is already scheduled for the selected week/time/court number.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "modify" then

				notUpdated = Request("notUpdated")

				If notUpdated = "" Then
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All Match Schedule information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontError10'><b>")
						rw(notUpdated)
						rw("</b></font>")
					rw("</td>")
					rw("</tr>")
				End If

				Session("admin") = ""

			ElseIf Session("Err") <> "" then
				rw(Session("Err"))
				Session("Err") = ""
			End If
		%>

		</table>

	<%
	End Sub

	'******************************************'

	Sub modifyRecord()

		WEEK = request("WEEK")


		''is WEEK selected yet?
		if WEEK <> "" then
rw("HERE")
Response.End
			sql = "SELECT w.WEEK_ID, w.WEEK_NUM, w.[DATE], e.EVENT_SHORT_DESC "_
				& "FROM WEEK_TBL w LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
				& "WHERE e.ACTIVE_EVENT_IND IN ('A','Y') "_
				& "AND w.WEEK_ID = '" & WEEK & "' "_
				& "ORDER BY UPPER(EVENT_SHORT_DESC), w.[DATE]"

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsWeekData = rs.GetRows
				rsWeekRows = UBound(rsWeekData,2)
			else
				rw("Error:Missing Weeks.<p />" & sql)
				Response.End
			end if

			sql = "SELECT e.EVENT_SHORT_DESC, t.MATCH_START_TIME, t.TIME_ID "_
				& "FROM WEEK_TBL w "_
				& "LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
				& "LEFT JOIN  TIME_TBL t "_
				& "ON e.EVENT_CD = t.EVENT_CD "_
				& "WHERE w.WEEK_ID = '" & WEEK & "' "_
				& "ORDER BY UPPER(e.EVENT_SHORT_DESC), t.MATCH_START_TIME"

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsTimeData = rs.GetRows
				rsTimeRows = UBound(rsTimeData,2)
			else
				rw("Error:Missing Times.<p />" & sql)
				Response.End
			end if

			''sql = "SELECT t.TEAM_ID, e.EVENT_SHORT_DESC + '/' + d.DIVISION_CD + '/' + t.TEAM_CD + '(' + t.TEAM_NAME + ')' as 'INFORMATION' "_
			sql = "SELECT t.TEAM_ID, e.EVENT_SHORT_DESC + '/' + d.DIVISION_CD + '/' + t.TEAM_CD as 'INFORMATION' "_
				& "FROM WEEK_TBL w "_
				& "LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
				& "LEFT JOIN DIVISION_TBL d ON e.EVENT_CD = d.EVENT_CD "_
				& "LEFT JOIN TEAM_TBL t ON d.DIVISION_ID = t.DIVISION_ID "_
				& "WHERE w.WEEK_ID = '" & WEEK & "' "_
				& "ORDER BY INFORMATION"

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsTeamData = rs.GetRows
				rsTeamRows = UBound(rsTeamData,2)
			else
				rw("Error:Missing Teams.<p />" & sql)
				Response.End
			end if






			''sql = "SELECT p.PERSON_ID, p.LAST_NAME + ', ' + p.FIRST_NAME + '(' + t.TEAM_CD + '/' + t.TEAM_NAME + ')' as 'PNAME' "_
			sql = "SELECT p.PERSON_ID, p.LAST_NAME + ', ' + p.FIRST_NAME + '(' + t.TEAM_CD + ')' as 'PNAME' "_
				& "FROM WEEK_TBL w  "_
				& "LEFT JOIN EVENT_TBL e ON w.EVENT_CD = e.EVENT_CD "_
				& "LEFT JOIN DIVISION_TBL d ON e.EVENT_CD = d.EVENT_CD "_
				& "LEFT JOIN TEAM_TBL t ON d.DIVISION_ID = t.DIVISION_ID "_
				& "LEFT JOIN TEAM_MEMBER_TBL tm ON t.TEAM_ID = tm.TEAM_ID "_
				& "LEFT JOIN db_accessadmin.PERSON_TBL p ON tm.PERSON_ID = p.PERSON_ID "_
				& "WHERE e.ACTIVE_EVENT_IND = 'Y' "_
				& "AND w.WEEK_ID = '" & WEEK & "' "_
				& "ORDER BY PNAME"

	''		sql = "SELECT p.PERSON_ID, p.LAST_NAME + ', ' + p.FIRST_NAME + '(' + t.TEAM_CD + '/' + t.TEAM_NAME + ')' as 'PNAME' "_
	''				& "FROM db_accessadmin.PERSON_TBL p "_
	''				& "LEFT JOIN REGISTRATION_TBL r ON p.PERSON_ID = r.PERSON_ID "_
	''				& "LEFT JOIN TEAM_MEMBER_TBL tm ON p.PERSON_ID = tm.PERSON_ID "_
	''				& "LEFT JOIN TEAM_TBL t ON tm.TEAM_ID = t.TEAM_ID "_
	''				& "LEFT JOIN EVENT_TBL e ON r.EVENT_CD = e.EVENT_CD "_
	''				& "LEFT JOIN WEEK_TBL w ON w.EVENT_CD = e.EVENT_CD "_
	''				& "WHERE e.ACTIVE_EVENT_IND = 'Y' "_
	''				& "AND w.WEEK_ID = '" & WEEK & "' "_
	''				& "AND r.REGISTRATION_IND = 'Y' "_
	''				& "ORDER BY PNAME"

			''rw(sql)
			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsPersonData = rs.GetRows
				rsPersonRows = UBound(rsPersonData,2)
			else
				rw("Error:Missing people registered for the chosen event.")
				Response.End
			end if

			SQL = "SELECT m.MATCH_ID, "_
				& "m.WEEK_ID, "_
				& "m.TIME_ID, "_
				& "m.COURT_NUM, "_
				& "m.TEAM1_TEAM_ID, "_
				& "m.TEAM2_TEAM_ID, "_
				& "m.REF_TEAM_ID, "_
				& "m.TEAM1_MVP_ID, "_
				& "m.TEAM2_MVP_ID, "_
				& "m.TEAM1_INCLUDE_DIV_STATS_IND, "_
				& "m.TEAM2_INCLUDE_DIV_STATS_IND "_
				& "FROM MATCH_SCHEDULE_TBL m "_
				& "WHERE m.WEEK_ID = '" & WEEK & "' "_
				& "ORDER BY m.COURT_NUM, m.TIME_ID"

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsMSData = rs.GetRows
				rsMSRows = UBound(rsMSData,2)
			else
				rsMSRows = -1
			end if

		else
			rsMSRows = -1
		end if


	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please select an event/week update the information.
		</b></font>
		</div>

		<br /><br />
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
		<input type='hidden' name='choice' value='modify' />
		</form>

		<p />

		<% if WEEK <> "" then%>
			<form name='MODIFY' method='post' action='MatchSchedule_submit.asp' onSubmit="return validateModify();">

			<table width='500' bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

			<tr bgcolor='#000066'>
			<th valign='bottom'><nobr><font class='cfont8'><a class='menuGray' href='javascript:checkAll(MODIFY.ID);'>Select All</a></nobr><br /><nobr><a class='menuGray' href='javascript:uncheckAll(MODIFY.ID);'>Clear All</a></nobr></font><br /><font class='cfontWhite10'><b>Modify</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Team 1</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Team 1 MVP</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Team 2</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Team 2 MVP</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Week</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Time</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Court #</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Ref Team</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Inc In Team 1 Stats?</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Inc In Team 2 Stats?</b></font></th>
			</tr>

			<%
				For i = 0 to rsMSRows

					ODD_ROW = NOT ODD_ROW

					If ODD_ROW Then
						BGCOLOR = "#FFFFFF"
					Else
						BGCOLOR = "#F0F0F0"
					End If
			%>

					<tr bgcolor="<%=BGCOLOR%>">
					<td align='center'><input type="checkbox" id="ID<%=rsMSData(0,i)%>" name="ID" value="<%=rsMSData(0,i)%>" onclick="disableCheck(this);"></td>

					<td>
					<select name='TEAM1_TEAM_ID' id="TEAM1_TEAM_ID<%=rsMSData(0,i)%>" disabled>

					<%
						For j = 0 to rsTeamRows
							rw("<option value='" & rsTeamData(0,j) & "'")

							If rsMSData(4,i) = rsTeamData(0,j) then
								rw(" selected")
							End If

							rw(">" & rsTeamData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td>
					<select name='TEAM1_MVP_ID' id="TEAM1_MVP_ID<%=rsMSData(0,i)%>" disabled>

					<%
						rw("<option value='0'>- select -</option>")

						For j = 0 to rsPersonRows
							rw("<option value='" & rsPersonData(0,j) & "'")

							If rsMSData(7,i) = rsPersonData(0,j) then
								rw(" selected")
							End If

							rw(">" & rsPersonData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td>
					<select name='TEAM2_TEAM_ID' id="TEAM2_TEAM_ID<%=rsMSData(0,i)%>" disabled>

					<%
						For j = 0 to rsTeamRows
							rw("<option value='" & rsTeamData(0,j) & "'")

							If rsMSData(5,i) = rsTeamData(0,j) then
								rw(" selected")
							End If

							rw(">" & rsTeamData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td>
					<select name='TEAM2_MVP_ID' id="TEAM2_MVP_ID<%=rsMSData(0,i)%>" disabled>

					<%
						rw("<option value='0'>- select -</option>")

						For j = 0 to rsPersonRows
							rw("<option value='" & rsPersonData(0,j) & "'")

							If rsMSData(8,i) = rsPersonData(0,j) then
								rw(" selected")
							End If

							rw(">" & rsPersonData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td>
					<select name='WEEK_ID' id="WEEK_ID<%=rsMSData(0,i)%>" disabled>

					<%
						For j = 0 to rsWeekRows
							rw("<option value='" & rsWeekData(0,j) & "'")

							If rsMSData(1,i) = rsWeekData(0,j) then
								rw(" selected")
							End If

							rw(">Week " & rsWeekData(1,j) & " - " & rsWeekData(2,j) & "</option>")
						Next
					%>

					</select>
					</td>
					<td>
					<select name='TIME_ID' id="TIME_ID<%=rsMSData(0,i)%>" disabled>

					<%
						For j = 0 to rsTimeRows
							rw("<option value='" & rsTimeData(2,j) & "'")

							If rsMSData(2,i) = rsTimeData(2,j) then
								rw(" selected")
							End If

							rw(">" & rsTimeData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td><input type="text" name='COURT_NUM' id="COURT_NUM<%=rsMSData(0,i)%>" size='2' maxlength='2' value='<%=rsMSData(3,i)%>' disabled /></td>

					<td>
					<select name='REF_TEAM_ID' id="REF_TEAM_ID<%=rsMSData(0,i)%>" disabled>

					<%
						For j = 0 to rsTeamRows
							rw("<option value='" & rsTeamData(0,j) & "'")

							If rsMSData(6,i) = rsTeamData(0,j) then
								rw(" selected")
							End If

							rw(">" & rsTeamData(1,j) & "</option>")
						Next
					%>

					</select>
					</td>

					<td>
					<select name='TEAM1_INCLUDE_DIV_STATS_IND' id="TEAM1_INCLUDE_DIV_STATS_IND<%=rsMSData(0,i)%>" disabled>
					<option value="Y" <%If rsMSData(9,i)="Y" then rw( "selected") End If%>>Yes</option>
					<option value="N" <%If rsMSData(9,i)="N" then rw( "selected") End If%>>No</option>
					</select>
					</td>

					<td>
					<select name='TEAM2_INCLUDE_DIV_STATS_IND' id="TEAM2_INCLUDE_DIV_STATS_IND<%=rsMSData(0,i)%>" disabled>
					<option value="Y" <%If rsMSData(10,i)="Y" then rw( "selected") End If%>>Yes</option>
					<option value="N" <%If rsMSData(10,i)="N" then rw( "selected") End If%>>No</option>
					</select>
					</td>

					</tr>
				<%Next%>
			<tr>
			<td colspan='2'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Modify Match Schedule(s)'></font></td>
			<td colspan='9' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Modify Match Schedule(s)'></font></td>
			</tr>

			</table>

			<input type='hidden' name='WEEK' value='<%=WEEK%>' />
			</form>
		<%end if%>



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

		<br /><br />


	<%
		call closeRSCNConnection()
	End Sub

	'******************************************'

	Sub insertRecord()

	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the new match schedule information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='MatchSchedule_submit.asp' onSubmit="return validateInsert();">

		<table bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Week</b></font></td>
		<td>
		<select name='WEEK_ID'>

		<%
			For i = 0 to rsWeekRows
				rw("<option value='" & rsWeekData(0,i) & "'")

				If Session("WEEK_ID") = CStr(rsWeekData(0,i)) then
					rw(" selected")
				End If

				rw(">" & rsWeekData(3,i) & " - Week " & rsWeekData(1,i) & " - " & rsWeekData(2,i) & "</option>")
			Next
		%>

		</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Time</b></font></td>
		<td>
		<select name='TIME_ID'>

		<%
			For i = 0 to rsTimeRows
				rw("<option value='" & rsTimeData(2,i) & "'")

				If Session("TIME_ID") = CStr(rsTimeData(2,i)) then
					rw(" selected")
				End If

				rw(">" & rsTimeData(0,i) & "/" & rsTimeData(1,i) & "</option>")
			Next
		%>

		</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Court Number</b></font></td>
		<td><input type="text" name='COURT_NUM' size='2' maxlength='2' value='<%=Session("COURT_NUM")%>' /></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 1</b></font></td>
		<td>
		<select name='TEAM1_TEAM_ID'>

		<%
			For i = 0 to rsTeamRows
				rw("<option value='" & rsTeamData(0,i) & "'")

				If Session("TEAM1_TEAM_ID") = CStr(rsTeamData(0,i)) then
					rw(" selected")
				End If

				rw(">" & rsTeamData(1,i) & "</option>")
			Next
		%>

		</select><font size='1' face='arial'> Event/Division/Team Code/(Team Name)</font>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 2</b></font></td>
		<td>
		<select name='TEAM2_TEAM_ID'>

		<%
			For i = 0 to rsTeamRows
				rw("<option value='" & rsTeamData(0,i) & "'")

				If Session("TEAM2_TEAM_ID") = CStr(rsTeamData(0,i)) then
					rw(" selected")
				End If

				rw(">" & rsTeamData(1,i) & "</option>")
			Next
		%>

		</select><font size='1' face='arial'> Event/Division/Team Code/(Team Name)</font>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Ref Team</b></font></td>
		<td>
		<select name='REF_TEAM_ID'>

		<%
			For i = 0 to rsTeamRows
				rw("<option value='" & rsTeamData(0,i) & "'")

				If Session("REF_TEAM_ID") = CStr(rsTeamData(0,i)) then
					rw(" selected")
				End If

				rw(">" & rsTeamData(1,i) & "</option>")
			Next
		%>

		</select><font size='1' face='arial'> Event/Division/Team Code/(Team Name)</font>
		</td>
		</tr>

		<!--
		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 1 MVP</b></font></td>
		<td>
		<select name='TEAM1_MVP_ID'>

		For i = 0 to
		rw(<option value=' & xxx(0,i) & '> & xxx(1,i) & </option>)
		Next

		</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 2 MVP</b></font></td>
		<td>
		<select name='TEAM2_MVP_ID'>

		For i = 0 to
		rw(<option value=' & xxx(0,i) & '> & xxx(1,i) & </option>)
		Next

		</select>
		</td>
		</tr>
		-->

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 1 Include In Division Stats?</b></font></td>
		<td>
		<select name='TEAM1_INCLUDE_DIV_STATS_IND'>
		<option value="Y" selected>Yes</option>
		<option value="N">No</option>
		</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Team 2 Include In Division Stats?</b></font></td>
		<td>
		<select name='TEAM2_INCLUDE_DIV_STATS_IND'>
		<option value="Y" selected>Yes</option>
		<option value="N">No</option>
		</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Match Schedule'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.WEEK_ID.focus();
		-->
		</script>
	<%

		closeRSCNConnection()
	End Sub
	%>
</td>
</tr>

<!--#include virtual="/incs/fragContact.asp"-->

</table>

<%


'******************************************'

Sub closePage()
	rw("</body>")
	rw("</html>")
End Sub

'******************************************'

%>



