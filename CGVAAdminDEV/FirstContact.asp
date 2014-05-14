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
<title>CGVA - First Contact Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{


	if(INSERT.FIRST_CONTACT_DESC.value == "")
	{
		alert("Please enter an first contact description.");
		INSERT.FIRST_CONTACT_DESC.focus();
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

				if(MODIFY.FIRST_CONTACT_DESC[selectedValue].value == "")
				{
					alert("Please enter a first contact description for the record being modified.");
					MODIFY.FIRST_CONTACT_DESC[selectedValue].focus();
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

		if(MODIFY.FIRST_CONTACT_DESC.value == "")
		{
			alert("Please enter a first contact description.");
			MODIFY.FIRST_CONTACT_DESC.focus();
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
	var FIRST_CONTACT_DESC = eval("MODIFY.FIRST_CONTACT_DESC" + rowValue);

	if(idValue.checked)
	{
		FIRST_CONTACT_DESC.disabled = false;
	}
	else
	{
		FIRST_CONTACT_DESC.disabled = true;
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
	<b><u>CGVA - First Contact Administration</u></b>
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

		<form name="CHOICE" method="POST" action="FirstContact.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New First Contact
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing First Contact
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
					rw("<b>The new event was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>An event with the same name already exists.</b>")
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
						rw("<b>All event information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontError10'>")
						rw("<b>Not All event information was modified successfully (duplicate record).</b>")
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

		SQL = "SELECT FIRST_CONTACT_ID, "_
					& "FIRST_CONTACT_DESC "_
			& "FROM 		FIRST_CONTACT_TBL "_
			& "ORDER BY 	UPPER(FIRST_CONTACT_DESC)"
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

		<form name='MODIFY' method='post' action='FirstContact_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='2'><input type='submit' name="submitChoice" value='Modify First Contact(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>First Contact Description</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("FIRST_CONTACT_ID")%>" name="ID" value="<%=rs("FIRST_CONTACT_ID")%>" onclick="disableCheck(this);"></td>
				<td><input type="text" id="FIRST_CONTACT_DESC<%=rs("FIRST_CONTACT_ID")%>" name="FIRST_CONTACT_DESC" value="<%=rs("FIRST_CONTACT_DESC")%>" size='45' maxlength='50' disabled></td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='2'><input type='submit' name="submitChoice" value='Modify First Contact(s)' onClick="return validateModify();"></td>
		</tr>

		</table>

		</form>

		<br /><br />


	<%
		''call closeRSCNConnection()
	End Sub

	'******************************************'

	Sub insertRecord()

	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the new first contact information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='FirstContact_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>First Contact Description</b></font></td>
		<td><input type="text" name='FIRST_CONTACT_DESC' size='45' maxlength='50'></td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add First Contact'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.FIRST_CONTACT_DESC.focus();
		-->
		</script>
	<%

		''closeRSCNConnection()
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


