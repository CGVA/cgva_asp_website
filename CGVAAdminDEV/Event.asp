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

	sql = "SELECT LOCATION_CD, LOCATION_DESC FROM LOCATION_TBL "_
		& "ORDER BY UPPER(LOCATION_DESC)"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsLocationData = rs.GetRows
		rsLocationRows = UBound(rsLocationData,2)
	else
		rw("Error:Missing Location Codes.")
		Response.End
	end if

	sql = "SELECT EVENT_TYPE_ID,EVENT_TYPE_DESC FROM EVENT_TYPE_TBL "_
		& "ORDER BY UPPER(EVENT_TYPE_DESC)"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsEventTypeData = rs.GetRows
		rsEventTypeRows = UBound(rsEventTypeData,2)
	else
		rw("Error:Missing Event Types.")
		Response.End
	end if

	sql = "SELECT WEEK_ID, EVENT_CD + ' - ' + CONVERT(varchar,WEEK_NUM) AS 'EVENT/WEEK#' FROM WEEK_TBL "_
		& "ORDER BY 'EVENT/WEEK#'"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsWNDData = rs.GetRows
		rsWNDRows = UBound(rsWNDData,2)
	else
		rsWNDRows = -1
	end if

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

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Event Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{

	if(INSERT.EVENT_CD.value == "")
	{
		alert("Please enter an event code description.");
		INSERT.EVENT_CD.focus();
		return false;
	}

	if(INSERT.EVENT_IMG.value == "")
	{
		alert("Please select an event graphic.");
		INSERT.EVENT_IMG.focus();
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

				if(MODIFY.EVENT_CD[selectedValue].value == "")
				{
					alert("Please enter an event code for the record being modified.");
					MODIFY.EVENT_CD[selectedValue].focus();
					return false;
				}

				if(MODIFY.EVENT_IMG[selectedValue].value == "")
				{
					alert("Please enter an event graphic for the record being modified.");
					MODIFY.EVENT_IMG[selectedValue].focus();
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

		if(MODIFY.EVENT_CD.value == "")
		{
			alert("Please enter an event code for the record being modified.");
			MODIFY.EVENT_CD.focus();
			return false;
		}

		if(MODIFY.EVENT_IMG.value == "")
		{
			alert("Please enter an event graphic for the record being modified.");
			MODIFY.EVENT_IMG.focus();
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
	var LOCATION_CD = eval("MODIFY.LOCATION_CD" + rowValue);
	var EVENT_SHORT_DESC = eval("MODIFY.EVENT_SHORT_DESC" + rowValue);
	var EVENT_LONG_DESC = eval("MODIFY.EVENT_LONG_DESC" + rowValue);
	var EVENT_START_DATE = eval("MODIFY.EVENT_START_DATE" + rowValue);
	var EVENT_END_DATE = eval("MODIFY.EVENT_END_DATE" + rowValue);
	var EVENT_TYPE_ID = eval("MODIFY.EVENT_TYPE_ID" + rowValue);
	var ACTIVE_EVENT_IND = eval("MODIFY.ACTIVE_EVENT_IND" + rowValue);
	var	WEEK_NUM_DISPLAY_IND = eval("MODIFY.WEEK_NUM_DISPLAY_IND" + rowValue);
	var	EVENT_IMG = eval("MODIFY.EVENT_IMG" + rowValue);

	if(idValue.checked)
	{
		EVENT_CD.disabled = false;
		LOCATION_CD.disabled = false;
		EVENT_SHORT_DESC.disabled = false;
		EVENT_LONG_DESC.disabled = false;
		EVENT_START_DATE.disabled = false;
		EVENT_END_DATE.disabled = false;
		EVENT_TYPE_ID.disabled = false;
		ACTIVE_EVENT_IND.disabled = false;
		WEEK_NUM_DISPLAY_IND.disabled = false;
		EVENT_IMG.disabled = false;
	}
	else
	{
		EVENT_CD.disabled = true;
		LOCATION_CD.disabled = true;
		EVENT_SHORT_DESC.disabled = true;
		EVENT_LONG_DESC.disabled = true;
		EVENT_START_DATE.disabled = true;
		EVENT_END_DATE.disabled = true;
		EVENT_TYPE_ID.disabled = true;
		ACTIVE_EVENT_IND.disabled = true;
		WEEK_NUM_DISPLAY_IND.disabled = true;
		EVENT_IMG.disabled = true;
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
	<b><u>CGVA - Event Administration</u></b>
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

		<form name="CHOICE" method="POST" action="Event.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Event
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Event
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
						rw("<font class='cfontSuccess10'>")
						rw("<b>All event information was modified successfully, <font class='cfontError10'>except</font> for these: " & notUpdated & ".</b>")
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

		SQL = "SELECT EVENT_CD, "_
					& "LOCATION_CD, "_
					& "EVENT_SHORT_DESC, "_
					& "EVENT_LONG_DESC, "_
					& "EVENT_START_DATE, "_
					& "EVENT_END_DATE, "_
					& "EVENT_TYPE_ID, "_
					& "ACTIVE_EVENT_IND, "_
					& "WEEK_NUM_DISPLAY_IND, "_
					& "EVENT_IMG "_
			& "FROM 		EVENT_TBL "_
			& "ORDER BY 	EVENT_CD"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

	%>

		<div align='center'>
		<font class='cfont10'><b>Select any record(s) to modify, and make any necessary changes.</b></font>
		<br />
		<font class='cfontError10'><b>*The event graphic should already be in the image table, and should be named without the _0, _1 etc. and should not contain the .gif at the end.*</b></font>
		<br /><br />

		<%
	''		If rs.EOF then
	''			rw("<font class='cfontError12'>* No data available to modify. *</font>")
	''			Response.End
	''		End If
		%>

		</div>

		<form name='MODIFY' method='post' action='Event_Submit.asp'>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Event(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Modify</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Location Code</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Short Desc</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Long Desc</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Start Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event End Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Type</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Event Graphic</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Approved Week # for website display</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Active?</b></font></th>
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
				<td align='center'><input type="checkbox" id="ID<%=rs("EVENT_CD")%>" name="ID" value="<%=rs("EVENT_CD")%>" onclick="disableCheck(this);"></td>
				<td><input type="text" id="EVENT_CD<%=rs("EVENT_CD")%>" name="EVENT_CD" value="<%=rs("EVENT_CD")%>" size='6' maxlength='6' disabled></td>

				<td>
					<select id="LOCATION_CD<%=rs("EVENT_CD")%>" name="LOCATION_CD" disabled>

					<%
						for i = 0 to rsLocationRows
							rw("<option value='" & rsLocationData(0,i) & "'")

							If rs("LOCATION_CD") = rsLocationData(0,i) then
								rw("selected")
							End If

							rw(">" & rsLocationData(1,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td><input type="text" id="EVENT_SHORT_DESC<%=rs("EVENT_CD")%>" name="EVENT_SHORT_DESC" value="<%=rs("EVENT_SHORT_DESC")%>" size='30' maxlength='40' disabled></td>
				<td><input type="text" id="EVENT_LONG_DESC<%=rs("EVENT_CD")%>" name="EVENT_LONG_DESC" value="<%=rs("EVENT_LONG_DESC")%>" size='40' maxlength='200' disabled></td>
				<td><input type="text" id="EVENT_START_DATE<%=rs("EVENT_CD")%>" name="EVENT_START_DATE" value="<%=rs("EVENT_START_DATE")%>" size='8' maxlength='10' disabled></td>
				<td><input type="text" id="EVENT_END_DATE<%=rs("EVENT_CD")%>" name="EVENT_END_DATE" value="<%=rs("EVENT_END_DATE")%>" size='8' maxlength='10' disabled></td>

				<td>
					<select id="EVENT_TYPE_ID<%=rs("EVENT_CD")%>" name="EVENT_TYPE_ID" disabled>

					<%
						for i = 0 to rsEventTypeRows
							rw("<option value='" & rsEventTypeData(0,i) & "'")

							If rs("EVENT_TYPE_ID") = rsEventTypeData(0,i) then
								rw("selected")
							End If

							rw(">" & rsEventTypeData(1,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td>
					<select id="EVENT_IMG<%=rs("EVENT_CD")%>" name="EVENT_IMG" disabled>
					<option value=''>-select-</option>

					<%
						for i = 0 to rsEventImageRows
							rw("<option value='" & rsEventImageData(0,i) & "'")

							If rs("EVENT_IMG") = rsEventImageData(0,i) then
								rw("selected")
							End If

							rw(">" & rsEventImageData(1,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td>
					<select id="WEEK_NUM_DISPLAY_IND<%=rs("EVENT_CD")%>" name="WEEK_NUM_DISPLAY_IND" disabled>
					<option value='0'>-none-</option>
					<%
						for i = 0 to rsWNDRows
							rw("<option value='" & rsWNDData(0,i) & "'")

							If rs("WEEK_NUM_DISPLAY_IND") = rsWNDData(0,i) then
								rw("selected")
							End If

							rw(">" & rsWNDData(1,i) & "</option>")
						next

					%>

					</select>
				</td>

				<td>
					<select id="ACTIVE_EVENT_IND<%=rs("EVENT_CD")%>" name="ACTIVE_EVENT_IND" disabled>

						<%
							rw("<option value='Y' ")

							If rs("ACTIVE_EVENT_IND") = "Y" then
								rw("selected")
							End If

							rw(">Yes</option>")

							rw("<option value='N' ")

							If rs("ACTIVE_EVENT_IND") = "N" then
								rw("selected")
							End If

							rw(">No</option>")

						%>

					</select>

				</td>

				</tr>

		<%
				rs.MoveNext
			Loop
		%>

		<tr>
		<td colspan='9'><input type='submit' name="submitChoice" value='Modify Event(s)' onClick="return validateModify();"></td>
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
		Please enter the new event information.
		</b></font>
		<br />
		<font class='cfontError10'><b>*The event graphic should already be in the image table, and should be named without the _0, _1 etc. and should not contain the .gif at the end.*</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Event_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Code</b></font></td>
		<td><input type="text" name="EVENT_CD" size='6' maxlength='6'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Location</b></font></td>
		<td>
			<select name="LOCATION_CD">

			<%
				For i = 0 to rsLocationRows
					rw("<option value='" & rsLocationData(0,i) & "'>" & rsLocationData(1,i) & "</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Short Description</b></font></td>
		<td><input type="text" name="EVENT_SHORT_DESC" size='30' maxlength='40'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Long Description</b></font></td>
		<td><input type="text" name="EVENT_LONG_DESC" size='50' maxlength='200'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Start Date</b></font></td>
		<td><input type="text" name="EVENT_START_DATE" size='20' maxlength='10'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event End Date</b></font></td>
		<td><input type="text" name="EVENT_END_DATE" size='20' maxlength='10'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Type</b></font></td>
		<td>
			<select name="EVENT_TYPE_ID">

			<%
				For i = 0 to rsEventTypeRows
					rw("<option value='" & rsEventTypeData(0,i) & "'>" & rsEventTypeData(1,i) & "</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event Graphic</b></font></td>
		<td>
			<select name="EVENT_IMG">
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
		<td align='right' valign='top'><font size='2' face='arial'><b>Approved Week #<br />for website display</b></font></td>
		<td>
			<select name="WEEK_NUM_DISPLAY_IND">
			<option value='0'>-none-</option>
			<%
				For i = 0 to rsWNDRows
					rw("<option value='" & rsWNDData(0,i) & "'>" & rsWNDData(1,i) & "</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Active?</b></font></td>
		<td>
			<select name="ACTIVE_EVENT_IND">
			<option value='Y'>Yes</option>
			<option value='N'>No</option>
			</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Event'></font></td>
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


