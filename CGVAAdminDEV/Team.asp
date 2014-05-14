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

	sql = "SELECT a.DIVISION_ID, b.EVENT_SHORT_DESC,a.DIVISION_DESC FROM DIVISION_TBL a "_
		& "LEFT JOIN EVENT_TBL b ON a.EVENT_CD = b.EVENT_CD "_
		& "WHERE b.ACTIVE_EVENT_IND = 'Y' "_
		& "ORDER BY UPPER(EVENT_SHORT_DESC), DIVISION_DESC"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsEvDivData = rs.GetRows
		rsEvDivRows = UBound(rsEvDivData,2)
	else
		rw("Error:Missing Active Events/Divisions.")
		Response.End
	end if

%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Team Administration</title>

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
	<b><u>CGVA - Team Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Team.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Team
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Team
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
					rw("<b>The new team was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A team with the same name/code already exists for the selected event/division.</b>")
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
						rw("<b>All team information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All team information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
						rw("</font>")
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

		SQL = "SELECT  a.TEAM_ID, "_
					& "a.DIVISION_ID, "_
					& "b.DIVISION_DESC, "_
					& "c.EVENT_SHORT_DESC, "_
					& "a.TEAM_CD, "_
					& "a.TEAM_NAME "_
			& "FROM TEAM_TBL a "_
			& "LEFT JOIN DIVISION_TBL b ON a.DIVISION_ID = b.DIVISION_ID "_
			& "LEFT JOIN EVENT_TBL c ON b.EVENT_CD = c.EVENT_CD "_
			& "WHERE c.ACTIVE_EVENT_IND= 'Y' "_
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

	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the new team information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Team_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event/Division</b></font></td>
		<td>
		<select name='DIVISION_ID'>

		<%
			For i = 0 to rsEvDivRows
				rw("<option value='" & rsEvDivData(0,i) & "'>" & rsEvDivData(1,i) & "/" & rsEvDivData(2,i) & "</option>")
			Next
		%>
		</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>TEAM_CD</b></font></td>
		<td><input type="text" name='TEAM_CD' size='3' maxlength='3'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>TEAM_NAME</b></font></td>
		<td><input type="text" name='TEAM_NAME' size='30' maxlength='50'></td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Team'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.DIVISION_ID.focus();
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


