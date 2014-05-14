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
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Event Type Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	if(INSERT.EVENT_TYPE_DESC.value == "")
	{
		alert("Please enter an event type description.");
		INSERT.EVENT_TYPE_DESC.focus();
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

				if(MODIFY.EVENT_TYPE_DESC[selectedValue].value == "")
				{
					alert("Please enter an event type description for the record being modified.");
					MODIFY.EVENT_TYPE_DESC[selectedValue].focus();
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

		if(MODIFY.EVENT_TYPE_DESC.value == "")
		{
			alert("Please enter an event type description for the record being modified.");
			MODIFY.EVENT_TYPE_DESC.focus();
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
	var EVENT_TYPE_DESC = eval("MODIFY.EVENT_TYPE_DESC" + rowValue);

	if(idValue.checked)
	{
		EVENT_TYPE_DESC.disabled = false;
	}
	else
	{
		EVENT_TYPE_DESC.disabled = true;
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
	<b><u>CGVA - Event Type Administration</u></b>
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

		<form name="CHOICE" method="POST" action="EventType.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Event Type
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Event Type
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
					rw("<b>The new event type was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>An event type with the same name already exists.</b>")
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
						rw("<b>All event type information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All event type information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
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

		SQL = "SELECT 		EVENT_TYPE_ID, "_
			& "				EVENT_TYPE_DESC "_
			& "FROM 		EVENT_TYPE_TBL "_
			& "ORDER BY 	EVENT_TYPE_DESC"
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

		<form name='MODIFY' method='post' action='EventType_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='11'><input type='submit' name="submitChoice" value='Modify Event Type(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Type Description</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("EVENT_TYPE_ID")%>" name="ID" value="<%=rs("EVENT_TYPE_ID")%>" onclick="disableCheck(this);"></td>
				<td><input type="text" id="EVENT_TYPE_DESC<%=rs("EVENT_TYPE_ID")%>" name="EVENT_TYPE_DESC" value="<%=rs("EVENT_TYPE_DESC")%>" size='17' maxlength='20' disabled></td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='11'><input type='submit' name="submitChoice" value='Modify Event Type(s)' onClick="return validateModify();"></td>
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
		Please enter the new event type information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='EventType_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Type Description</b></font></td>
		<td><input type="text" name="EVENT_TYPE_DESC" size='30' maxlength='20'></td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Event Type'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.EVENT_TYPE_DESC.focus();
		-->
		</script>
	<%

		closeCNConnection()
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


