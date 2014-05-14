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

'	sql = "SELECT a.DIVISION_ID, b.EVENT_SHORT_DESC,a.DIVISION_DESC FROM DIVISION_TBL a "_
'		& "LEFT JOIN EVENT_TBL b ON a.EVENT_CD = b.EVENT_CD "_
'		& "WHERE b.ACTIVE_EVENT_IND = 'Y' "_
'		& "ORDER BY UPPER(EVENT_SHORT_DESC), DIVISION_DESC"
'
'	set rs = cn.Execute(sql)
'
'	if not rs.EOF then
'		rsEvDivData = rs.GetRows
'		rsEvDivRows = UBound(rsEvDivData,2)
'	else
'		rw("Error:Missing Active Events/Divisions.")
'		Response.End
'	end if

%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Team Member Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	if(INSERT.TEAM_CD.value == "")
	{
		alert("Please enter a team code.");
		INSERT.TEAM_CD.focus();
		return false;
	}


	if(INSERT.TEAM_NAME.value == "")
	{
		alert("Please enter a team name.");
		INSERT.TEAM_NAME.focus();
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

				if(MODIFY.TEAM_CD[selectedValue].value == "")
				{
					alert("Please enter a team code for the record being modified.");
					MODIFY.TEAM_CD[selectedValue].focus();
					return false;
				}

				if(MODIFY.TEAM_NAME[selectedValue].value == "")
				{
					alert("Please enter a team name for the record being modified.");
					MODIFY.TEAM_NAME[selectedValue].focus();
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


		if(MODIFY.TEAM_CD.value == "")
		{
			alert("Please enter a team code for the record being modified.");
			MODIFY.TEAM_CD.focus();
			return false;
		}

		if(MODIFY.TEAM_NAME.value == "")
		{
			alert("Please enter a team name for the record being modified.");
			MODIFY.TEAM_NAME.focus();
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
	//var TEAM_ID = eval("MODIFY.TEAM_ID" + rowValue);
	var DIVISION_ID = eval("MODIFY.DIVISION_ID" + rowValue);
	var TEAM_CD = eval("MODIFY.TEAM_CD" + rowValue);
	var TEAM_NAME = eval("MODIFY.TEAM_NAME" + rowValue);

	if(idValue.checked)
	{

		//TEAM_ID.disabled = false;
		DIVISION_ID.disabled = false;
		TEAM_CD.disabled = false;
		TEAM_NAME.disabled = false;
	}else{
		//TEAM_ID.disabled = true;
		DIVISION_ID.disabled = true;
		TEAM_CD.disabled = true;
		TEAM_NAME.disabled = true;
	}

}




//-->

</script>

</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Team Member Administration</u></b>
	</font>
	</div>

	<br />

	<%
		'call subroutines'
''		If Request("choice") = "" then
''			Call chooseOption()
''			Call closePage()
''		ElseIf Request("choice") = "modify" then
''			Call modifyRecord()
''			Call closePage()
''		ElseIf Request("choice") = "insert" then
			Call insertRecord()
			Call closePage()
''		End If


	'******************************************'

	Sub chooseOption()
	%>

		<form name="CHOICE" method="POST" action="TeamMember.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Manage Team Rosters
			</font>
		</td>
		</tr>

		<!--
		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> View Team Rosters
			</font>
		</td>
		</tr>
		-->

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

				notUpdated = Request("notUpdated")

				If notUpdated = "" Then
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All team member information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontError10'>")
						rw("<b>You tried to enter a person on this team that already is on another team for this event (IDs: " & notupdated & ").</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				End If

			End If
			Session("admin") = ""
			Session("Err") = ""
		%>

		</table>

	<%
	End Sub

	'******************************************'

	Sub modifyRecord()

		SQL = "SELECT  a.TEAM_ID, "_
					& "a.DIVISION_ID, "_
					& "b.DIVISION_DESC, "_
					& "c.EVENT_SHORT_DESC, "_
					& "a.TEAM_CD, "_
					& "a.TEAM_NAME "_
			& "FROM TEAM_TBL a "_
			& "LEFT JOIN DIVISION_TBL b ON a.DIVISION_ID = b.DIVISION_ID "_
			& "LEFT JOIN EVENT_TBL c ON b.EVENT_CD = c.EVENT_CD "_
			& "ORDER BY 	UPPER(EVENT_SHORT_DESC), DIVISION_DESC"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

	%>

		<div align='center'>
		<font class='cfont10'><b>Select any record(s) to modify, and make any necessary changes.</b></font>
		<br /><br />

		<%
	''		If rs.EOF then
	''			rw("<font class='cfontError12'>* No data available to modify. *</font>")
	''			Response.End
	''		End If
		%>

		</div>

		<form name='MODIFY' method='post' action='Team_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Team(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event/Division</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Team Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Team Description</b></font></th>
		</tr>

		<%
			Do While not rs.EOF


				ODD_ROW = NOT ODD_ROW

				If ODD_ROW Then
					BGCOLOR = "#FFFFFF"
				Else
					BGCOLOR = "#F0F0F0"
				End If
		%>

				<tr bgcolor="<%=BGCOLOR%>">
				<td align='center'><input type="checkbox" id="ID<%=rs("TEAM_ID")%>" name="ID" value="<%=rs("TEAM_ID")%>" onclick="disableCheck(this);"></td>

				<td>
					<select id="DIVISION_ID<%=rs("TEAM_ID")%>" name="DIVISION_ID" disabled>

					<%
						for i = 0 to rsEvDivRows
							rw("<option value='" & rsEvDivData(0,i) & "'")

							If rs("DIVISION_ID")  =  rsEvDivData(0,i) then
								rw("selected")
							End If

							rw(">" & rsEvDivData(1,i) & "/" & rsEvDivData(2,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td><input type="text" id="TEAM_CD<%=rs("TEAM_ID")%>" name="TEAM_CD" value="<%=rs("TEAM_CD")%>" size='3' maxlength='3' disabled></td>
				<td><input type="text" id="TEAM_NAME<%=rs("TEAM_ID")%>" name="TEAM_NAME" value="<%=rs("TEAM_NAME")%>" size='30' maxlength='50' disabled></td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Team(s)' onClick="return validateModify();"></td>
		</tr>

		</table>

		</form>

		<br /><br />


	<%
		call closeRSCNConnection()
	End Sub

	'******************************************'

	Sub insertRecord()
		EDT = request("EDT")

		sql = "SELECT b.EVENT_SHORT_DESC, a.DIVISION_DESC, "_
			& "t.TEAM_ID, t.TEAM_CD, T.TEAM_NAME, b.EVENT_CD "_
			& "FROM TEAM_TBL t "_
			& "LEFT JOIN DIVISION_TBL a ON t.DIVISION_ID = a.DIVISION_ID "_
			& "LEFT JOIN EVENT_TBL b ON a.EVENT_CD = b.EVENT_CD "_
			& "WHERE b.ACTIVE_EVENT_IND = 'Y' "_
			& "ORDER BY UPPER(EVENT_SHORT_DESC), DIVISION_DESC, TEAM_CD"

		''rw(sql)
		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsEvDivData = rs.GetRows
			rsEvDivRows = UBound(rsEvDivData,2)
		else
			rw("Error:Missing Active Events/Divisions/Teams.")
			Response.End
		end if

		''is team selected yet?
		if EDT <> "" then
			EDTArray = Split(EDT,"_")
			EVENT_CD = EDTArray(0)
			TEAM = EDTArray(1)

			SQL = "SELECT t.PERSON_ID, CAPTAIN_IND, CERTIFIED_REF_IND "_
				& "FROM TEAM_MEMBER_TBL t "_
				& "LEFT JOIN db_accessadmin.PERSON_TBL p ON t.PERSON_ID = p.PERSON_ID "_
				& "WHERE TEAM_ID = '" & TEAM & "' "_
				& "ORDER BY UPPER(p.LAST_NAME), UPPER(p.FIRST_NAME)"

			''rw(sql & "<br />")
			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsTeamMemberData = rs.GetRows
				rsTeamMemberRows = UBound(rsTeamMemberData,2)
			else
				rsTeamMemberRows = -1
			end if

			sql = "SELECT a.PERSON_ID, b.FIRST_NAME, b.LAST_NAME "_
				& "FROM REGISTRATION_TBL a "_
				& "LEFT JOIN db_accessadmin.PERSON_TBL b ON a.PERSON_ID = b.PERSON_ID "_
				& "WHERE a.EVENT_CD = '" & EVENT_CD & "' "_
				& "AND a.REGISTRATION_IND = 'Y' "_
				& "ORDER BY UPPER(b.LAST_NAME), UPPER(b.FIRST_NAME)"

			''& "AND a.PERSON_ID NOT IN (SELECT PERSON_ID FROM TEAM_MEMBER_TBL) "_
			''rw(sql & "<br />")
			''Response.End

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsPlayerData = rs.GetRows
				rsPlayerRows = UBound(rsPlayerData,2)
			else
				rw("Error:Missing Player Records.")
				Response.End
			end if

			sql = "SELECT a.PERSON_ID, b.FIRST_NAME, b.LAST_NAME "_
				& "FROM REGISTRATION_TBL a "_
				& "LEFT JOIN db_accessadmin.PERSON_TBL b ON a.PERSON_ID = b.PERSON_ID "_
				& "WHERE a.EVENT_CD = '" & EVENT_CD & "' "_
				& "AND a.REGISTRATION_IND = 'Y' "_
				& "AND a.PERSON_ID NOT IN (SELECT PERSON_ID FROM TEAM_MEMBER_TBL tm " _
				& "							LEFT JOIN TEAM_TBL t ON tm.TEAM_ID = t.TEAM_ID " _
				& "							LEFT JOIN DIVISION_TBL d ON t.DIVISION_ID = d.DIVISION_ID " _
				& "							WHERE d.EVENT_CD = '" & EVENT_CD & "') "_
				& "ORDER BY UPPER(b.LAST_NAME), UPPER(b.FIRST_NAME)"

			''rw(sql & "<br />")
			''Response.End

			set rs = cn.Execute(sql)

			if not rs.EOF then
				rsPlayer2Data = rs.GetRows
				rsPlayer2Rows = UBound(rsPlayer2Data,2)
			else
				rsPlayer2Rows = -1
			end if

		else
			rsTeamMemberRows = -1
		end if


	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please select an event/team and enter/update the new team member information.
		</b></font>
		</div>

		<br /><br />
		<form name='teamChoice'>
		<div align='center'>
			<font class='cfont10'>
			Event/Division/Team:
			<select name='EDT' onChange="changePage();">
			<option value=''>-select-</option>
			<%
				For i = 0 to rsEvDivRows
					rw("<option value='" & rsEvDivData(5,i) & "_" & rsEvDivData(2,i) & "'")

					If EDT  =  rsEvDivData(5,i) & "_" & rsEvDivData(2,i) then
						rw(" selected")
					End If

					rw(">" & rsEvDivData(0,i) & " - " & rsEvDivData(1,i) & " - " & rsEvDivData(3,i) & " (" & rsEvDivData(4,i) & ")</option>")
				Next
			%>
			</select>
			</font>
		</div>
		<input type='hidden' name='choice' value='insert' />
		</form>

		<p />

		<% if EDT <> "" then%>
			<form name='INSERT' method='post' action='TeamMember_submit.asp' onSubmit="return validateInsert();">

			<table width='500' bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

			<tr bgcolor='#000066'>
			<th valign='bottom'><font class='cfontWhite10'><b>&nbsp;</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Name</b></font></th>
			<th valign='bottom'><font class='cfontWhite10'><b>Captain?</b></font></th>
			<!--<th valign='bottom'><font class='cfontWhite10'><b>Certified Ref?</b></font></th>-->
			</tr>

			<%
				For i = 0 to rsTeamMemberRows


					ODD_ROW = NOT ODD_ROW

					If ODD_ROW Then
						BGCOLOR = "#FFFFFF"
					Else
						BGCOLOR = "#F0F0F0"
					End If
			%>

					<tr bgcolor="<%=BGCOLOR%>">
					<td align='right'><font class='cfont10'><b><%=i+1%>.</b></font></td>
					<td>
						<select name='PERSON_ID'>
						<option value=''>-select-</option>
						<%
							for j = 0 to rsPlayerRows
								rw("<option value='" & rsPlayerData(0,j) & "'")

								If rsTeamMemberData(0,i) = rsPlayerData(0,j) then
									rw(" selected")
								End If

								rw(">" & rsPlayerData(2,j) & ", " & rsPlayerData(1,j) & "</option>")
							Next
						%>
						</select>
					</td>

					<td>
						<select name='CAPTAIN_IND'>
						<option value='N'<%If rsTeamMemberData(1,i)="N" Then rw(" selected") End If%>>No</option>
						<option value='Y'<%If rsTeamMemberData(1,i)="Y" Then rw(" selected") End If%>>Yes</option>
						</select>
					</td>

					<!--
					<td>
						<select name='CERTIFIED_REF_IND'>
						<option value='N'<'%'If rsTeamMemberData(2,i)="N" Then rw(" selected") End If'%'>>No</option>
						<option value='Y'<'%'If rsTeamMemberData(2,i)="Y" Then rw(" selected") End If'%'>>Yes</option>
						</select>
					</td>
					-->
					</tr>

			<%
				Next
				''reset counter if no members currently exist for a team
				''display up to 10 total rows per team for adding new members
				''If rsTeamMemberRows = -1 Then rsTeamMemberRows = 0 End If

				''rw(rsTeamMemberRows)
				For i = rsTeamMemberRows+1 to 9
					ODD_ROW = NOT ODD_ROW

					If ODD_ROW Then
						BGCOLOR = "#FFFFFF"
					Else
						BGCOLOR = "#F0F0F0"
					End If
			%>
					<tr bgcolor="<%=BGCOLOR%>">
					<td align='right'><font class='cfont10'><b><%=i+1%>.</b></font></td>
					<td>
						<select name='PERSON_ID'>
						<option value=''>-select-</option>
						<%
							for j = 0 to rsPlayer2Rows
								rw("<option value='" & rsPlayer2Data(0,j) & "'>" & rsPlayer2Data(2,j) & ", " & rsPlayer2Data(1,j) & "</option>")
							Next
						%>
						</select>
					</td>

					<td>
						<select name='CAPTAIN_IND'>
						<option value='N'>No</option>
						<option value='Y'>Yes</option>
						</select>
					</td>

					<!--
					<td>
						<select name='CERTIFIED_REF_IND'>
						<option value='N'>No</option>
						<option value='Y'>Yes</option>
						</select>
					</td>
					-->
					</tr>
			<%
				Next
			%>

			<tr>
			<td colspan='3' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Save Team Members'></font></td>
			</tr>

			</table>

			<input type='hidden' name='EDT' value='<%=EDT%>' />
			</form>
		<%end if%>



		<script language='Javascript'>
		<!--

			function changePage()
			{
				choice = teamChoice.EDT.options[teamChoice.EDT.selectedIndex].value;
				//alert(choice);
				if(choice != "")
				{
					teamChoice.submit();
				}
			}
		//-->
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



<!--
function disableCheck(obj)
{

rowValue = obj.value;
var idValue = eval(MODIFY.ID + rowValue);
var PERSON_ID = eval(MODIFY.PERSON_ID + rowValue);
var TEAM_ID = eval(MODIFY.TEAM_ID + rowValue);
var CAPTAIN_IND = eval(MODIFY.CAPTAIN_IND + rowValue);
//var CERTIFIED_REF_IND = eval(MODIFY.CERTIFIED_REF_IND + rowValue);
if(idValue.checked)
{

PERSON_ID.disabled = false;
TEAM_ID.disabled = false;
CAPTAIN_IND.disabled = false;
//CERTIFIED_REF_IND.disabled = false;
}else{

PERSON_ID.disabled = true;
TEAM_ID.disabled = true;
CAPTAIN_IND.disabled = true;
//CERTIFIED_REF_IND.disabled = true;
}

}
<tr>
<td align='right' valign='middle'><font size='2' face='arial'><b>PERSON_ID</b></font></td>
<td>
<select name='PERSON_ID'>
<option value=''>-select-</option>
for j = 0 to
	rw(<option value=' & rsPlayerData(0,j) & ')

	If rsTeamMemberData(PLAYER ID HERE x,i)  =  rsPlayerData(0,j) then
		rw("selected")
	End If

	rw(> & rsPlayerData(2,j) & ", " & rsPlayerData(1,j) & </option>)
Next

</select>
</td>
</tr>

<input type="hidden" name='TEAM_ID' value='<TEAM>' />

<tr>
<td align='right' valign='middle'><font size='2' face='arial'><b>CAPTAIN_IND</b></font></td>
<td>
	<select name='CAPTAIN_IND'>
	<option value='N'>No</option>
	<option value='Y'>Yes</option>
	</select>
</td>
</tr>

<tr>
<td align='right' valign='middle'><font size='2' face='arial'><b>CERTIFIED_REF_IND</b></font></td>
<td><input type="text" name='CERTIFIED_REF_IND' size='' maxlength=''></td>
</tr>

<tr>
<td align='right' valign='middle'><font size='2' face='arial'><b>CERTIFIED_REF_IND</b></font></td>
<td>
	<select name='CERTIFIED_REF_IND'>
	<option value='N'>No</option>
	<option value='Y'>Yes</option>
	</select>
</td>
</tr>

-->

