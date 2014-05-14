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

	sql = "SELECT PERSON_ID, FIRST_NAME,LAST_NAME FROM db_accessadmin.PERSON_TBL "_
		& "ORDER BY UPPER(LAST_NAME), UPPER(FIRST_NAME)"
'' "WHERE PERSON_ID NOT IN (SELECT PERSON_ID FROM RATER_TBL) "_

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsRaterData = rs.GetRows
		rsRaterRows = UBound(rsRaterData,2)
	else
		rw("Error:Missing Raters.")
		Response.End
	end if


%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Rater Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	/*
	if(INSERT.EVENT_CD.value == "")
	{
		alert("Please enter an event code description.");
		INSERT.EVENT_CD.focus();
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
		if(MODIFY.EVENT_CD[selectedValue].value == "")
		{
			alert("Please enter an event code for the record being modified.");
			MODIFY.EVENT_CD[selectedValue].focus();
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


		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID.focus();
			return false;
		}

		/*
		if(MODIFY.EVENT_CD.value == "")
		{
			alert("Please enter an event code for the record being modified.");
			MODIFY.EVENT_CD.focus();
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
	var RATER = eval("MODIFY.RATER" + rowValue);

	if(idValue.checked)
	{
		RATER.disabled = false;
	}
	else
	{
		RATER.disabled = true;
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
	<b><u>CGVA - Rater Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Rater.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Rater
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Rater
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
					rw("<b>The new rater was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A rater with the same name already exists.</b>")
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
						rw("<b>All rater information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontError10'>")
						rw("<b>Not All rater information was modified successfully (duplicate record).</b>")
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

		SQL = "SELECT 		a.PERSON_ID, b.FIRST_NAME, b.LAST_NAME "_
			& "FROM 		RATER_TBL a LEFT JOIN db_accessadmin.PERSON_TBL b on a.PERSON_ID = b.PERSON_ID "_
			& "ORDER BY 	UPPER(b.LAST_NAME), UPPER(b.FIRST_NAME)"
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

		<form name='MODIFY' method='post' action='Rater_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='2'><input type='submit' name="submitChoice" value='Modify Rater(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Rater</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("PERSON_ID")%>" name="ID" value="<%=rs("PERSON_ID")%>" onclick="disableCheck(this);"></td>
				<td>
					<select id="RATER<%=rs("PERSON_ID")%>" name="RATER" disabled>

					<%
						for i = 0 to rsRaterRows
							rw("<option value='" & rsRaterData(0,i) & "'")

							If rs("PERSON_ID") = rsRaterData(0,i) then
								rw("selected")
							End If

							rw(">" & rsRaterData(2,i) & ", " & rsRaterData(1,i) & "(PID: " & rsRaterData(0,i) & ")</option>")
						next

					%>

					</select>
				</td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='2'><input type='submit' name="submitChoice" value='Modify Rater(s)' onClick="return validateModify();"></td>
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
		Please enter the new rater information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Rater_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Rater</b></font></td>
		<td>
			<select name='RATER'>
			<%
				For i = 0 to rsRaterRows
					rw("<option value='" & rsRaterData(0,i) & "'>" & rsRaterData(2,i) & ", " & rsRaterData(1,i) & "(PID: " & rsRaterData(0,i) & ")</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Rater'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.RATER.focus();
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


