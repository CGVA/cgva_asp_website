<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	If Session("USER_ID") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("index.asp")
	End If
%>

<html>

<head>
<title>CGVA - Personnel Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	/*
	if(INSERTUSER.role.value == "")
	{
		alert("Please select a role.");
		INSERTUSER.role.focus();
		return false;
	}

	if(INSERTUSER.loginID.value == "")
	{
		alert("Please enter this user's Siteminder login ID.");
		INSERTUSER.loginID.focus();
		return false;
	}
	*/

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
			}

		}

		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID[0].focus();
			return false;
		}

		/*
		if(MODIFY.loginID[selectedValue].value == "")
		{
			alert("Please enter a login ID for the record being modified.");
			MODIFY.loginID[selectedValue].focus();
			return false;
		}
		*/

	}
	else
	{

		if(MODIFY.ID.checked)
		{
			selectedValue = 1;
		}


		/*
		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID.focus();
			return false;
		}

		if(MODIFY.loginID.value == "")
		{
			alert("Please enter a login ID for the record being modified.");
			MODIFY.loginID.focus();
			return false;
		}
		*/

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
	var role = eval("MODIFY.role" + rowValue);
	var login = eval("MODIFY.loginID" + rowValue);
	var active = eval("MODIFY.active" + rowValue);


	if(idValue.checked)
	{
		role.disabled = false;
		login.disabled = false;
		active.disabled = false;
	}
	else
	{
		role.disabled = true;
		login.disabled = true;
		active.disabled = true;
	}

}

//-->

</script>

</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->


<div align='center'>
<font class='cfont12'>
<b><u>CGVA - Personnel Administration</u></b>
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

	<form name="CHOICE" method="POST" action="Personnel.asp">

	<table align='center' border='0' cellpadding='3'>

	<tr>
	<td>
		<font class='cfont10'>
		<input type="radio" name="choice" value="insert" checked> Add New CGVA Member
		</font>
	</td>
	</tr>

	<tr>
	<td>
		<font class='cfont10'>
		<input type="radio" name="choice" value="modify"> Modify Existing CGVA Member
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
				rw("<b>The new user was added successfully.</b>")
				rw("</font>")
			rw("</td>")
			rw("</tr>")

			Session("admin") = ""

		ElseIf Session("admin") = "insertFail" then
			rw("<tr>")
			rw("<td>")
				rw("<font class='cfontError10'>")
				rw("<b>A person with the same first/last name and email address/phone you entered already exists.</b>")
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
					rw("<b>All user information was modified successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")
			Else
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontSuccess10'>")
					rw("<b>All user information was modified successfully, <font class='cfontError10'>except</font> for these login IDs (can't change login ID to one that already exists within SDAT): " & notUpdated & ".</b>")
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

''	SQL = "SELECT 		a.USER_ID, "_
''		& "				LOGIN_ID, "_
''		& "				FIRST_NAME, "_
''		& "				LAST_NAME, "_
''		& "				ROLE_ID, "_
''		& "				ACTIVE_FLAG "_
''		& "FROM 		T_SDAT_LOGIN a "_
''		& "LEFT JOIN 	T_SDAT_ACCESS b "_
''		& "ON			a.USER_ID = b.USER_ID "_
''		& "WHERE 		APP_ID = 8 "_
''		& "ORDER BY 	ROLE_ID, ACTIVE_FLAG DESC, UPPER(LOGIN_ID)"
''	rw("<!-- SQL: " & SQL & " -->")
''	Set rs = cn.Execute(SQL)

''	SQL = "SELECT * FROM T_REF_SDAT_ROLE ORDER BY ROLE_ID"
''	Set rs2 = cn.Execute(SQL)

''	alldata = rs2.getrows
''	numrows = ubound(alldata,2)

''	rs2.close
''	set rs2 = Nothing
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

	<form name='MODIFY' method='post' action='Personnel_Submit.asp'>


	<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

	<tr>
	<td colspan='7'><input type='submit' name="submitChoice" value='Modify User(s)' onClick="return validateModify();"></td>
	</tr>

	<tr bgcolor='#000066'>
	<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
	<th valign='bottom'><font class='cfontWhite10'><b>Role</b></font></th>
	<th valign='bottom'><font class='cfontWhite10'><b>Login ID</b></font></th>
	<th valign='bottom'><font class='cfontWhite10'><b>Active?</b></font></th>
	</tr>


	<tr>
	<td colspan='7'><input type='submit' name="submitChoice" value='Modify User(s)' onClick="return validateModify();"></td>
	</tr>

	</table>

	</form>

	<br /><br />


<%
	call closeRSCNConnection()
End Sub

'******************************************'

Sub insertRecord()

''	Select Case Session("roleID")
''
''	Case 0:
''		SQL = "SELECT * FROM T_REF_SDAT_ROLE ORDER BY ROLE_ID"
''		Set rs = cn.Execute(SQL)
''	Case 1:
''		SQL = "SELECT * FROM T_REF_SDAT_ROLE WHERE ROLE_ID > 0 ORDER BY ROLE_ID"
''		Set rs = cn.Execute(SQL)
''	End Select

%>

	<div align='center'>
	<font class='cfont10'><b>
	Please enter the new user information.
	</b></font>
	</div>

	<br /><br />

	<form name='INSERTUSER' method='post' action='Personnel_submit.asp' onSubmit="return validateInsert();">

	<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

	<tr>
	<td align='right' valign='middle'><font size='2' face='arial'><b>Role</b></font></td>
	<td>
		<select name="role">
		<option value="">-select-</option>

		<%
''			Do While not rs.EOF
''				rw("<option value='" & rs("ROLE_ID") & "'>" & rs("ROLE_DESC") & "</option>")
''				rs.MoveNext
''			Loop
		%>
		</select>
	</td>
	</tr>

	<tr>
	<td align='right' valign='middle'><font size='2' face='arial'><b>Login ID</b></font></td>
	<td><input type="text" name="loginID" maxlength='50'></td>
	</tr>

	<tr>
	<td align='right' valign='middle'><font size='2' face='arial'><b>Active?</b></font></td>
	<td>
		<select name="active">
		<option value="Y" selected>Yes</option>
		<option value="N">No</option>
		</select>
	</td>
	</tr>

	<tr>
	<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add User'></font></td>
	</tr>

	</table>

	</form>

	<script>
	<!--
		INSERTUSER.role.focus();
	-->
	</script>
<%

	closeRSCNConnection()
End Sub


'******************************************'

Sub closePage()
	rw("</body>")
	rw("</html>")
End Sub

'******************************************'

%>


