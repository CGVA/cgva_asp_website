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
		Response.Redirect("index.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	sql = "SELECT EVENT_CD,EVENT_SHORT_DESC FROM EVENT_TBL "_
		& "WHERE ACTIVE_EVENT_IND IN ('Y','A') "_
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
<title>CGVA - Division Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{


	if(INSERT.DIVISION_CD.value == "")
	{
		alert("Please enter a division code.");
		INSERT.DIVISION_CD.focus();
		return false;
	}

	if(INSERT.DIVISION_DESC.value == "")
	{
		alert("Please enter a division description.");
		INSERT.DIVISION_DESC.focus();
		return false;
	}

	if(INSERT.DIVISION_IMG.value == "")
	{
		alert("Please enter a division graphic.");
		INSERT.DIVISION_IMG.focus();
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

				if(MODIFY.DIVISION_IMG[selectedValue].value == "")
				{
					alert("Please enter a division graphic for the record being modified.");
					MODIFY.DIVISION_IMG[selectedValue].focus();
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

		if(MODIFY.DIVISION_IMG.value == "")
		{
			alert("Please enter a division graphic for the record being modified.");
			MODIFY.DIVISION_IMG.focus();
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
	var EVENT_CD = eval("MODIFY.EVENT_CD" + rowValue);
	var DIVISION_CD = eval("MODIFY.DIVISION_CD" + rowValue);
	var DIVISION_DESC = eval("MODIFY.DIVISION_DESC" + rowValue);
	var DIVISION_IMG = eval("MODIFY.DIVISION_IMG" + rowValue);

	if(idValue.checked)
	{
		EVENT_CD.disabled = false;
		DIVISION_CD.disabled = false;
		DIVISION_DESC.disabled = false;
		DIVISION_IMG.disabled = false;
	}
	else
	{
		EVENT_CD.disabled = true;
		DIVISION_CD.disabled = true;
		DIVISION_DESC.disabled = true;
		DIVISION_IMG.disabled = true;
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
	<b><u>CGVA - Division Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Division.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Division
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Division
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
					rw("<b>The new division was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A division with the same name already exists for the selected event.</b>")
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
						rw("<b>All division information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All division information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
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

		sql = "SELECT BUTTON_REF_STRING, BUTTON_FILENAME FROM BUTTON_REF_TBL "_
			& "ORDER BY BUTTON_REF_STRING"

		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsEventImageData = rs.GetRows
			rsEventImageRows = UBound(rsEventImageData,2)
		else
			rsImageRows = -1
		end if

		SQL = "SELECT DIVISION_ID, "_
					& "d.EVENT_CD, "_
					& "DIVISION_CD, "_
					& "DIVISION_DESC, "_
					& "DIVISION_IMG "_
			& "FROM 		DIVISION_TBL d "_
			& "LEFT JOIN 		EVENT_TBL e "_
			& "ON d.EVENT_CD = e.EVENT_CD "_
			& "WHERE e.ACTIVE_EVENT_IND IN ('Y','A') "_
			& "ORDER BY 	d.EVENT_CD, UPPER(DIVISION_DESC)"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

	%>

		<div align='center'>
		<font class='cfont10'><b>Select any record(s) to modify, and make any necessary changes.</b></font>
		<br />
		<font class='cfontError10'><b>*The division graphic should already be in the image table, and should be named without the _on, _off etc. and should not contain the .gif at the end.*</b></font>
		<br /><br />

		<%
	''		If rs.EOF then
	''			rw("<font class='cfontError12'>* No data available to modify. *</font>")
	''			Response.End
	''		End If
		%>

		</div>

		<form name='MODIFY' method='post' action='Division_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Division(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Division Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Division Description</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Division Graphic</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("DIVISION_ID")%>" name="ID" value="<%=rs("DIVISION_ID")%>" onclick="disableCheck(this);"></td>

				<td>
					<select id="EVENT_CD<%=rs("DIVISION_ID")%>" name="EVENT_CD" disabled>

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

				<td><input type="text" id="DIVISION_CD<%=rs("DIVISION_ID")%>" name="DIVISION_CD" value="<%=rs("DIVISION_CD")%>" size='6' maxlength='6' disabled></td>
				<td><input type="text" id="DIVISION_DESC<%=rs("DIVISION_ID")%>" name="DIVISION_DESC" value="<%=rs("DIVISION_DESC")%>" size='20' maxlength='25' disabled></td>
				<td>
					<select id="DIVISION_IMG<%=rs("DIVISION_ID")%>" name="DIVISION_IMG" disabled>
					<option value=''>-select</option>
					<%
						for i = 0 to rsEventImageRows
							rw("<option value='" & rsEventImageData(0,i) & "'")

							If rs("DIVISION_IMG") = rsEventImageData(0,i) then
								rw("selected")
							End If

							rw(">" & rsEventImageData(1,i) & "</option>")
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
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Division(s)' onClick="return validateModify();"></td>
		</tr>

		</table>

		</form>

		<br /><br />


	<%
		call closeRSCNConnection()
	End Sub

	'******************************************'

	Sub insertRecord()

		sql = "SELECT BUTTON_REF_STRING, BUTTON_FILENAME FROM BUTTON_REF_TBL "_
			& "ORDER BY BUTTON_REF_STRING"

		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsEventImageData = rs.GetRows
			rsEventImageRows = UBound(rsEventImageData,2)
		else
			rsImageRows = -1
		end if
	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the new division information.
		</b></font>
		<br />
		<font class='cfontError10'><b>*The division graphic should already be in the image table, and should be named without the _on, _off etc. and should not contain the .gif at the end.*</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Division_submit.asp' onSubmit="return validateInsert();">

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
		<td align='right' valign='middle'><font size='2' face='arial'><b>Division Code</b></font></td>
		<td><input type="text" name='DIVISION_CD' size='6' maxlength='6'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Division Description</b></font></td>
		<td><input type="text" name='DIVISION_DESC' size='20' maxlength='25'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Division Graphic</b></font></td>
		<td>
			<select name='DIVISION_IMG'>
			<option value=''>-select-</option>
			<%
				For i = 0 to rsEventImageRows
					rw("<option value='" & rsEventImageData(0,i) & "'>" & rsEventImageData(1,i) & "</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Division'></font></td>
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


