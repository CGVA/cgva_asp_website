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
<title>CGVA - Location Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	if(INSERT.LOCATION_CD.value == "")
	{
		alert("Please enter a location code.");
		INSERT.LOCATION_CD.focus();
		return false;
	}

	if(INSERT.LOCATION_DESC.value == "")
	{
		alert("Please enter a location description.");
		INSERT.LOCATION_DESC.focus();
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

				if(MODIFY.LOCATION_CD[selectedValue].value == "")
				{
					alert("Please enter a location code for the record being modified.");
					MODIFY.LOCATION_CD[selectedValue].focus();
					return false;
				}

				if(MODIFY.LOCATION_DESC[selectedValue].value == "")
				{
					alert("Please enter a location description for the record being modified.");
					MODIFY.LOCATION_DESC[selectedValue].focus();
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

		if(MODIFY.LOCATION_CD.checked)
		{
			selectedValue = 1;
		}


		if(selectedValue == "x")
		{
			alert("Please select a record to modify.");
			MODIFY.ID.focus();
			return false;
		}

		if(MODIFY.LOCATION_CD.value == "")
		{
			alert("Please enter a location code for the record being modified.");
			MODIFY.LOCATION_CD.focus();
			return false;
		}

		if(MODIFY.LOCATION_DESC.value == "")
		{
			alert("Please enter a location description for the record being modified.");
			MODIFY.LOCATION_DESC.focus();
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
	var LOCATION_CD = eval("MODIFY.LOCATION_CD" + rowValue);
	var LOCATION_DESC = eval("MODIFY.LOCATION_DESC" + rowValue);
	var LOCATION_PHONE = eval("MODIFY.LOCATION_PHONE" + rowValue);
	var LOCATION_FIRST_NAME = eval("MODIFY.LOCATION_FIRST_NAME" + rowValue);
	var LOCATION_LAST_NAME = eval("MODIFY.LOCATION_LAST_NAME" + rowValue);
	var ADDRESS_LINE1 = eval("MODIFY.ADDRESS_LINE1" + rowValue);
	var ADDRESS_LINE2 = eval("MODIFY.ADDRESS_LINE2" + rowValue);
	var CITY = eval("MODIFY.CITY" + rowValue);
	var STATE = eval("MODIFY.STATE" + rowValue);
	var ZIP = eval("MODIFY.ZIP" + rowValue);


	if(idValue.checked)
	{
		LOCATION_CD.disabled = false;
		LOCATION_DESC.disabled = false;
		LOCATION_PHONE.disabled = false;
		LOCATION_FIRST_NAME.disabled = false;
		LOCATION_LAST_NAME.disabled = false;
		ADDRESS_LINE1.disabled = false;
		ADDRESS_LINE2.disabled = false;
		CITY.disabled = false;
		STATE.disabled = false;
		ZIP.disabled = false;
	}
	else
	{
		LOCATION_CD.disabled = true;
		LOCATION_DESC.disabled = true;
		LOCATION_PHONE.disabled = true;
		LOCATION_FIRST_NAME.disabled = true;
		LOCATION_LAST_NAME.disabled = true;
		ADDRESS_LINE1.disabled = true;
		ADDRESS_LINE2.disabled = true;
		CITY.disabled = true;
		STATE.disabled = true;
		ZIP.disabled = true;
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
	<b><u>CGVA - Location Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Location.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Location
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Location(s)
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
					rw("<b>The new location was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A location with the same name already exists.</b>")
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
						rw("<b>All location information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All location information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
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

		SQL = "SELECT 		LOCATION_CD, "_
			& "				LOCATION_DESC, "_
			& "				LOCATION_PHONE, "_
			& "				LOCATION_FIRST_NAME, "_
			& "				LOCATION_LAST_NAME, "_
			& "				ADDRESS_LINE1, "_
			& "				ADDRESS_LINE2, "_
			& "				CITY, "_
			& "				STATE, "_
			& "				ZIP "_
			& "FROM 		LOCATION_TBL "_
			& "ORDER BY 	LOCATION_DESC"
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

		<form name='MODIFY' method='post' action='Location_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='11'><input type='submit' name="submitChoice" value='Modify Location(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location Description</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location Phone</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location First Name</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location Last Name</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Address Line 1</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Address Line 2</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>City</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>State</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Zip Code</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("LOCATION_CD")%>" name="ID" value="<%=rs("LOCATION_CD")%>" onclick="disableCheck(this);"></td>
				<td><input type="text" id="LOCATION_CD<%=rs("LOCATION_CD")%>" name="LOCATION_CD" value="<%=rs("LOCATION_CD")%>" size='3' maxlength='2' disabled></td>
				<td><input type="text" id="LOCATION_DESC<%=rs("LOCATION_CD")%>" name="LOCATION_DESC" value="<%=rs("LOCATION_DESC")%>" size='25' maxlength='30' disabled></td>
				<td><input type="text" id="LOCATION_PHONE<%=rs("LOCATION_CD")%>" name="LOCATION_PHONE" value="<%=rs("LOCATION_PHONE")%>" size='11' maxlength='10' disabled></td>
				<td><input type="text" id="LOCATION_FIRST_NAME<%=rs("LOCATION_CD")%>" name="LOCATION_FIRST_NAME" value="<%=rs("LOCATION_FIRST_NAME")%>" size='20' maxlength='25' disabled></td>
				<td><input type="text" id="LOCATION_LAST_NAME<%=rs("LOCATION_CD")%>" name="LOCATION_LAST_NAME" value="<%=rs("LOCATION_LAST_NAME")%>" size='30' maxlength='35' disabled></td>
				<td><input type="text" id="ADDRESS_LINE1<%=rs("LOCATION_CD")%>" name="ADDRESS_LINE1" value="<%=rs("ADDRESS_LINE1")%>" size='35' maxlength='40' disabled></td>
				<td><input type="text" id="ADDRESS_LINE2<%=rs("LOCATION_CD")%>" name="ADDRESS_LINE2" value="<%=rs("ADDRESS_LINE2")%>" size='35' maxlength='40' disabled></td>
				<td><input type="text" id="CITY<%=rs("LOCATION_CD")%>" name="CITY" value="<%=rs("CITY")%>" size='35' maxlength='30' disabled></td>
				<td><input type="text" id="STATE<%=rs("LOCATION_CD")%>" name="STATE" value="<%=rs("STATE")%>" size='3' maxlength='2' disabled></td>
				<td><input type="text" id="ZIP<%=rs("LOCATION_CD")%>" name="ZIP" value="<%=rs("ZIP")%>" size='11' maxlength='10' disabled></td>
				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='11'><input type='submit' name="submitChoice" value='Modify Location(s)' onClick="return validateModify();"></td>
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
		Please enter the new location information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Location_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location Code</b></font></td>
		<td><input type="text" name="LOCATION_CD" size='3' maxlength='2'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location Description</b></font></td>
		<td><input type="text" name="LOCATION_DESC" size='40' maxlength='30'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location Phone</td>
		<td><input type="text" name="LOCATION_PHONE" size='11' maxlength='10'><font size='2' face='arial'><b>(numbers only)</b></font></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location First Name</b></font></td>
		<td><input type="text" name="LOCATION_FIRST_NAME" size='40' maxlength='25'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location Last Name</b></font></td>
		<td><input type="text" name="LOCATION_LAST_NAME" size='40' maxlength='35'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Address Line 1</b></font></td>
		<td><input type="text" name="ADDRESS_LINE1" size='40' maxlength='40'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Address Line 2</b></font></td>
		<td><input type="text" name="ADDRESS_LINE2" size='40' maxlength='40'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>City</b></font></td>
		<td><input type="text" name="CITY" size='40' maxlength='30'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>State</b></font></td>
		<td><input type="text" name="STATE" size='3' maxlength='2'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Zip</b></font></td>
		<td><input type="text" name="ZIP" size='11' maxlength='10'></td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Location'></font></td>
		</tr>

		</table>

		</form>

		<script>
		<!--
			INSERT.LOCATION_CD.focus();
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


