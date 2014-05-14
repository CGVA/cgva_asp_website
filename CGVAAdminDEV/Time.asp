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
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	sql = "SELECT EVENT_CD,EVENT_SHORT_DESC FROM EVENT_TBL "_
		& "WHERE ACTIVE_EVENT_IND = 'Y' "_
		& "ORDER BY UPPER(EVENT_SHORT_DESC)"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsEventData = rs.GetRows
		rsEventRows = UBound(rsEventData,2)
	else
		rw("Error:Missing Events.")
		Response.End
	end if

%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Time Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	if(INSERT.MATCH_NUM.value == "")
	{
		alert("Please enter a match number.");
		INSERT.MATCH_NUM.focus();
		return false;
	}

	if(INSERT.MATCH_START_TIME.value == "")
	{
		alert("Please enter a match start time ('06:10PM').");
		INSERT.MATCH_START_TIME.focus();
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

				if(MODIFY.MATCH_NUM[selectedValue].value == "")
				{
					alert("Please enter a match number for the record being modified.");
					MODIFY.MATCH_NUM[selectedValue].focus();
					return false;
				}

				if(MODIFY.MATCH_START_TIME[selectedValue].value == "")
				{
					alert("Please enter a start time for the record being modified ('06:10PM').");
					MODIFY.MATCH_START_TIME[selectedValue].focus();
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


		if(MODIFY.MATCH_NUM.value == "")
		{
			alert("Please enter a match number for the record being modified.");
			MODIFY.MATCH_NUM.focus();
			return false;
		}

		if(MODIFY.MATCH_START_TIME.value == "")
		{
			alert("Please enter a start time for the record being modified.");
			MODIFY.MATCH_START_TIME.focus();
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
	//var TIME_ID = eval("MODIFY.TIME_ID" + rowValue);
	var EVENT_CD = eval("MODIFY.EVENT_CD" + rowValue);
	var MATCH_NUM = eval("MODIFY.MATCH_NUM" + rowValue);
	var MATCH_START_TIME = eval("MODIFY.MATCH_START_TIME" + rowValue);

	if(idValue.checked)
	{
		//TIME_ID.disabled = false;
		EVENT_CD.disabled = false;
		MATCH_NUM.disabled = false;
		MATCH_START_TIME.disabled = false;
	}else{

		//TIME_ID.disabled = true;
		EVENT_CD.disabled = true;
		MATCH_NUM.disabled = true;
		MATCH_START_TIME.disabled = true;
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
	<b><u>CGVA - Time Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Time.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Time
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Time
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
					rw("<b>The new time was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A time with the same name already exists for the selected event.</b>")
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
						rw("<b>All time information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All time information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
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

		SQL = "SELECT TIME_ID, "_
					& "a.EVENT_CD, "_
					& "MATCH_NUM, "_
					& "MATCH_START_TIME "_
			& "FROM TIME_TBL a "_
			& "LEFT JOIN EVENT_TBL b ON a.EVENT_CD = b.EVENT_CD "_
			& "WHERE b.ACTIVE_EVENT_IND = 'Y' "_
			& "ORDER BY UPPER(b.EVENT_SHORT_DESC), MATCH_NUM"
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

		<form name='MODIFY' method='post' action='Time_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='4'><input type='submit' name="submitChoice" value='Modify Time(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Match Number</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Start Time</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("TIME_ID")%>" name="ID" value="<%=rs("TIME_ID")%>" onclick="disableCheck(this);"></td>

				<td>
					<select id="EVENT_CD<%=rs("TIME_ID")%>" name="EVENT_CD" disabled>

					<%
						for i = 0 to rsEventRows
							rw("<option value='" & rsEventData(0,i) & "'")

							If rs("EVENT_CD") = rsEventData(0,i) then
								rw("selected")
							End If

							rw(">" & rsEventData(1,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td><input type="text" id="MATCH_NUM<%=rs("TIME_ID")%>" name="MATCH_NUM" value="<%=rs("MATCH_NUM")%>" size='2' maxlength='2' disabled></td>
				<td><input type="text" id="MATCH_START_TIME<%=rs("TIME_ID")%>" name="MATCH_START_TIME" value="<%=rs("MATCH_START_TIME")%>" size='7' maxlength='7' disabled></td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='4'><input type='submit' name="submitChoice" value='Modify Time(s)' onClick="return validateModify();"></td>
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
		Please enter the new time information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Time_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Code</b></font></td>
		<td>
			<select name='EVENT_CD'>
			<%
				For i = 0 to rsEventRows
					rw("<option value='" & rsEventData(0,i) & "'>" & rsEventData(1,i) & "</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Match Number</b></font></td>
		<td><input type="text" name='MATCH_NUM' size='2' maxlength='2'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Match Start Time</b></font></td>
		<td><input type="text" name='MATCH_START_TIME' size='7' maxlength='7'></td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Time'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.EVENT_CD.focus();
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



