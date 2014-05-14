<%@ Language=VBScript %>
<%Response.Buffer = False%>
<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	''On Error Resume Next
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
		& "WHERE ACTIVE_EVENT_IND IN ('Y','A') "_
		& "ORDER BY EVENT_END_DATE DESC"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsEventData = rs.GetRows
		rsEventRows = UBound(rsEventData,2)
	else
		rw("Error:Missing Events.")
		Response.End
	end if

	sql = "SELECT PERSON_ID, "_
		& "FIRST_NAME, "_
		& "LAST_NAME "_
		& "FROM 		db_accessadmin.PERSON_TBL "_
		& "ORDER BY 	UPPER(LAST_NAME),UPPER(FIRST_NAME)"
	rw("<!-- SQL: " & sql & " -->")
	Set rs = cn.Execute(sql)

	if not rs.EOF then
		rsPersonData = rs.GetRows
		rsPersonRows = UBound(rsPersonData,2)
	else
		rw("Error:Missing Personnel.")
		Response.End
	end if
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Registration Administration</title>

<script language='JavaScript'>
<!--

function validateInsert()
{


	if(INSERT.OPEN_PLAY_IND.value == "N" && INSERT.REGISTRATION_IND.value == "N")
	{
		alert("Both the open play/registration indicators cannot be 'No'.");
		INSERT.OPEN_PLAY_IND.focus();
		return false;
	}

	if(INSERT.OPEN_PLAY_IND.value == "Y" && INSERT.REGISTRATION_IND.value == "Y")
	{
		alert("Both the open play/registration indicators cannot be 'Yes'.");
		INSERT.OPEN_PLAY_IND.focus();
		return false;
	}

	if(INSERT.DATE.value == "")
	{
		alert("Please enter a valid date.");
		INSERT.DATE.focus();
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

//			if(MODIFY.ID[i].checked)
//			{
				selectedValue = i;

				if(MODIFY.OPEN_PLAY_IND[selectedValue].value == "N" && MODIFY.REGISTRATION_IND[selectedValue].value == "N")
				{
					alert("Both the open play/registration indicators cannot be 'No'.");
					MODIFY.OPEN_PLAY_IND[selectedValue].focus();
					return false;
				}

				if(MODIFY.OPEN_PLAY_IND[selectedValue].value == "Y" && MODIFY.REGISTRATION_IND[selectedValue].value == "Y")
				{
					alert("Both the open play/registration indicators cannot be 'Yes'.");
					MODIFY.OPEN_PLAY_IND[selectedValue].focus();
					return false;
				}

				if(MODIFY.DATE[selectedValue].value == "")
				{
					alert("Please enter a valid date.");
					MODIFY.DATE[selectedValue].focus();
					return false;
				}

//			}

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

//		if(MODIFY.ID.checked)
//		{
//			selectedValue = 1;
//		}


//		if(selectedValue == "x")
//		{
//			alert("Please select a record to modify.");
//			MODIFY.ID.focus();
//			return false;
//		}

		if(MODIFY.OPEN_PLAY_IND.value == "N" && MODIFY.REGISTRATION_IND.value == "N")
		{
			alert("Both the open play/registration indicators cannot be 'No'.");
			MODIFY.OPEN_PLAY_IND.focus();
			return false;
		}

		if(MODIFY.OPEN_PLAY_IND.value == "Y" && MODIFY.REGISTRATION_IND.value == "Y")
		{
			alert("Both the open play/registration indicators cannot be 'Yes'.");
			MODIFY.OPEN_PLAY_IND.focus();
			return false;
		}

		if(MODIFY.DATE.value == "")
		{
			alert("Please enter a valid date.");
			MODIFY.DATE.focus();
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
	var PERSON_ID = eval("MODIFY.PERSON_ID" + rowValue);
	var vDATE = eval("MODIFY.DATE" + rowValue);
	var WAIVER_IND = eval("MODIFY.WAIVER_IND" + rowValue);
	var OPEN_PLAY_IND = eval("MODIFY.OPEN_PLAY_IND" + rowValue);
	var PLAY_UPDOWN_IND = eval("MODIFY.PLAY_UPDOWN_IND" + rowValue);
	var REGISTRATION_IND = eval("MODIFY.REGISTRATION_IND" + rowValue);
	var DOLLARS_COLLECTED = eval("MODIFY.DOLLARS_COLLECTED" + rowValue);
	var DOLLARS_OFF_COUPON = eval("MODIFY.DOLLARS_OFF_COUPON" + rowValue);
	var CHECK_AMT_COLLECTED = eval("MODIFY.CHECK_AMT_COLLECTED" + rowValue);
	var CHECK_NUM = eval("MODIFY.CHECK_NUM" + rowValue);
	var NOTES = eval("MODIFY.NOTES" + rowValue);

	if(idValue.checked)
	{
		EVENT_CD.disabled = false;
		PERSON_ID.disabled = false;
		vDATE.disabled = false;
		WAIVER_IND.disabled = false;
		OPEN_PLAY_IND.disabled = false;
		PLAY_UPDOWN_IND.disabled = false;
		REGISTRATION_IND.disabled = false;
		DOLLARS_COLLECTED.disabled = false;
		DOLLARS_OFF_COUPON.disabled = false;
		CHECK_AMT_COLLECTED.disabled = false;
		CHECK_NUM.disabled = false;
		NOTES.disabled = false;
	}
	else
	{
		EVENT_CD.disabled = true;
		PERSON_ID.disabled = true;
		vDATE.disabled = true;
		WAIVER_IND.disabled = true;
		OPEN_PLAY_IND.disabled = true;
		PLAY_UPDOWN_IND.disabled = true;
		REGISTRATION_IND.disabled = true;
		DOLLARS_COLLECTED.disabled = true;
		DOLLARS_OFF_COUPON.disabled = true;
		CHECK_AMT_COLLECTED.disabled = true;
		CHECK_NUM.disabled = true;
		NOTES.disabled = true;
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
	<b><u>CGVA - Registration Administration</u></b>
	</font>
	</div>

	<br />

	<%
		'call subroutines'
		If Request("choice") = "" then
			Call chooseOption()
			''Call closePage()
		ElseIf Request("choice") = "modify" then
			Call modifyRecord()
			''Call closePage()
		ElseIf Request("choice") = "insert" then
			Call insertRecord()
			''Call closePage()
		End If


	'******************************************'

	Sub chooseOption()
	%>

		<form name="CHOICE" method="POST" action="Registration.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Registration
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Registration
			&nbsp;
			<select name='FILTER'>
			<option value='A'>A</option>
			<option value='B'>B</option>
			<option value='C'>C</option>
			<option value='D'>D</option>
			<option value='E'>E</option>
			<option value='F'>F</option>
			<option value='G'>G</option>
			<option value='H'>H</option>
			<option value='I'>I</option>
			<option value='J'>J</option>
			<option value='K'>K</option>
			<option value='L'>L</option>
			<option value='M'>M</option>
			<option value='N'>N</option>
			<option value='O'>O</option>
			<option value='P'>P</option>
			<option value='Q'>Q</option>
			<option value='R'>R</option>
			<option value='S'>S</option>
			<option value='T'>T</option>
			<option value='U'>U</option>
			<option value='V'>V</option>
			<option value='W'>W</option>
			<option value='X'>X</option>
			<option value='Y'>Y</option>
			<option value='Z'>Z</option>
			<option value=''>All</option>
			</select>
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
					rw("<b>The new registration was added successfully.</b>")
					rw("</font>")
				rw("</td>")
				rw("</tr>")

				Session("admin") = ""

			ElseIf Session("admin") = "insertFail" then
				rw("<tr>")
				rw("<td>")
					rw("<font class='cfontError10'>")
					rw("<b>A registration with the same name/event/open play/registration/date already exists, or this person has already been registered for this event.</b>")
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
						rw("<b>All registration information was modified successfully.</b>")
						rw("</font>")
					rw("</td>")
					rw("</tr>")
				Else
					rw("<tr>")
					rw("<td>")
						rw("<font class='cfontSuccess10'>")
						rw("<b>All registration information was modified successfully, <font class='cfontError10'>except</font> for these (duplicates existed): " & notUpdated & ".</b>")
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

		vFILTER = Request("FILTER")

		SQL = "SELECT r.ID, "_
			& "r.EVENT_CD, "_
			& "r.PERSON_ID, "_
			& "p.LAST_NAME, "_
			& "p.FIRST_NAME, "_
			& "Replace(CONVERT(varchar(12),r.[DATE],101),'01/01/1900','') as '[DATE]', "_
			& "r.WAIVER_IND, "_
			& "r.OPEN_PLAY_IND, "_
			& "r.PLAY_UPDOWN_IND, "_
			& "r.REGISTRATION_IND, "_
			& "r.DOLLARS_COLLECTED, "_
			& "r.DOLLARS_OFF_COUPON, "_
			& "r.CHECK_AMT_COLLECTED, "_
			& "r.CHECK_NUM, "_
			& "r.NOTES, "_
			& "r.CAPTAIN_INTEREST_IND, "_
			& "r.VOLUNTEER_INTEREST_IND "_
			& "FROM 		REGISTRATION_TBL r "_
			& "LEFT JOIN	db_accessadmin.PERSON_TBL p ON r.PERSON_ID = p.PERSON_ID "_
			& "LEFT JOIN	EVENT_TBL e ON r.EVENT_CD = e.EVENT_CD "_
			& "WHERE LAST_NAME LIKE '" & vFILTER & "%' "_
			& "AND e.ACTIVE_EVENT_IND IN ('Y','A') "_
			& "ORDER BY 	EVENT_CD, UPPER(LAST_NAME),UPPER(FIRST_NAME)"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

	%>



		<div align='center'>
		<font class='cfont10'><b>Select any record(s) to modify, and make any necessary changes.</b></font>
		<br />
		<font class='cfontError10'>Please do not use commas in the notes field!</font>
		<br /><br />

		<%
	''		If rs.EOF then
	''			rw("<font class='cfontError12'>* No data available to modify. *</font>")
	''			Response.End
	''		End If
		%>

		</div>

		<form name='MODIFY' method='post' action='Registration_Submit.asp'>

		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='14'><input type='submit' name="submitChoice" value='Modify Registration(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Event</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Person</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Waiver?</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Open Play?</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Registered?</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>$ Collected</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>$ Off Coupon</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Check Amt Collected</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Check Num</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Notes</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Volunteer Interest?</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Captain Interest?</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Play Up/Down?</b></font></th>
		</tr>

		<%
			counter = 1
			Do While not rs.EOF


				ODD_ROW = NOT ODD_ROW

				If ODD_ROW Then
					BGCOLOR = "#FFFFFF"
				Else
					BGCOLOR = "#F0F0F0"
				End If
		%>

				<tr bgcolor="<%=BGCOLOR%>">
				<td valign='top'>
					<input type="hidden" id="ID<%=rs("ID")%>" name="ID" value="<%=rs("ID")%>" />
					<select id="EVENT_CD<%=rs("ID")%>" name="EVENT_CD">

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
				<td valign='top'>
					<select id="PERSON_ID<%=rs("ID")%>" name="PERSON_ID">

					<%
						for i = 0 to rsPersonRows
							rw("<option value='" & rsPersonData(0,i) & "'")

							If rs("PERSON_ID") = rsPersonData(0,i) then
								rw("selected")
							End If

							rw(">" & rsPersonData(2,i) & ", " & rsPersonData(1,i) & "</option>")
						next

					%>

					</select>
				</td>
				<td valign='top'><input type="text" id="DATE<%=rs("ID")%>" name="DATE" value="<%=rs("[DATE]")%>" size='10' maxlength='10'></td>
				<td valign='top'>
					<select id="WAIVER_IND<%=rs("ID")%>" name="WAIVER_IND">
					<%
						rw("<option value='N' ")

						If rs("WAIVER_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='Y' ")

						If rs("WAIVER_IND") = "Y" then
							rw("selected")
						End If

						rw(">Yes</option>")

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="OPEN_PLAY_IND<%=rs("ID")%>" name="OPEN_PLAY_IND">
					<%
						rw("<option value='N' ")

						If rs("OPEN_PLAY_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='Y' ")

						If rs("OPEN_PLAY_IND") = "Y" then
							rw("selected")
						End If

						rw(">Yes</option>")

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="REGISTRATION_IND<%=rs("ID")%>" name="REGISTRATION_IND">
					<%
						rw("<option value='N' ")

						If rs("REGISTRATION_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='Y' ")

						If rs("REGISTRATION_IND") = "Y" then
							rw("selected")
						End If

						rw(">Yes</option>")

					%>

					</select>
				</td>
				<td valign='top'>$<input type="text" id="DOLLARS_COLLECTED<%=rs("ID")%>" name="DOLLARS_COLLECTED" value="<%=rs("DOLLARS_COLLECTED")%>" size='5' maxlength='25'></td>
				<td valign='top'>$<input type="text" id="DOLLARS_OFF_COUPON<%=rs("ID")%>" name="DOLLARS_OFF_COUPON" value="<%=rs("DOLLARS_OFF_COUPON")%>" size='5' maxlength='25'></td>
				<td valign='top'><input type="text" id="CHECK_AMT_COLLECTED<%=rs("ID")%>" name="CHECK_AMT_COLLECTED" value="<%=rs("CHECK_AMT_COLLECTED")%>" size='5' maxlength='25'></td>
				<td valign='top'><input type="text" id="CHECK_NUM<%=rs("ID")%>" name="CHECK_NUM" value="<%=rs("CHECK_NUM")%>" size='5' maxlength='25'></td>
				<td valign='top' align='center'><textarea id="NOTES<%=rs("ID")%>" name="NOTES"><%=rs("NOTES")%></textarea></td>
				<td valign='top'>
					<select id="VOLUNTEER_INTEREST_IND<%=rs("ID")%>" name="VOLUNTEER_INTEREST_IND">
					<%
						rw("<option value='N' ")

						If rs("VOLUNTEER_INTEREST_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='Y' ")

						If rs("VOLUNTEER_INTEREST_IND") = "Y" then
							rw("selected")
						End If

						rw(">Yes</option>")

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="CAPTAIN_INTEREST_IND<%=rs("ID")%>" name="CAPTAIN_INTEREST_IND">
					<%
						rw("<option value='N' ")

						If rs("CAPTAIN_INTEREST_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='Y' ")

						If rs("CAPTAIN_INTEREST_IND") = "Y" then
							rw("selected")
						End If

						rw(">Yes</option>")

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="PLAY_UPDOWN_IND<%=rs("ID")%>" name="PLAY_UPDOWN_IND">
					<%
						rw("<option value='N' ")

						If rs("PLAY_UPDOWN_IND") = "N" then
							rw("selected")
						End If

						rw(">No</option>")

						rw("<option value='D' ")

						If rs("PLAY_UPDOWN_IND") = "D" then
							rw("selected")
						End If

						rw(">Down</option>")

						rw("<option value='U' ")

						If rs("PLAY_UPDOWN_IND") = "U" then
							rw("selected")
						End If

						rw(">Up</option>")

						rw("<option value='UD' ")

						If rs("PLAY_UPDOWN_IND") = "UD" then
							rw("selected")
						End If

						rw(">Up or Down</option>")

					%>

					</select>
				</td>
    			</tr>

		<%
				rs.MoveNext
				counter = counter + 1
				response.flush
			Loop
		%>

		<tr>
		<td colspan='14'><input type='submit' name="submitChoice" value='Modify Registration(s)' onClick="return validateModify();"></td>
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
		Please enter the new registration information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='Registration_submit.asp' onSubmit="return validateInsert();">

		<table width='600' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Event</b></font></td>
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
		<td align='right' valign='middle'><font size='2' face='arial'><b>Person</b></font></td>
		<td>
			<select name='PERSON_ID'>

			<%
				For i = 0 to rsPersonRows
					rw("<option value='" & rsPersonData(0,i) & "'>" & rsPersonData(2,i) & ", " & rsPersonData(1,i) & "</option>")
				Next
			%>

			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Date</b></font></td>
		<td><input type="text" name='DATE' size='25' maxlength='25'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Waiver Collected?</b></font></td>
		<td>
			<select name='WAIVER_IND'>
			<option value='N'>No</option>
			<option value='Y' selected>Yes</option>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Open Play?</b></font></td>
		<td>
			<select name='OPEN_PLAY_IND'>
			<option value='N'>No</option>
			<option value='Y'>Yes</option>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Full Registration?</b></font></td>
		<td>
			<select name='REGISTRATION_IND'>
			<option value='N'>No</option>
			<option value='Y'>Yes</option>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Cash Collected</b></font></td>
		<td>$<input type="text" name='DOLLARS_COLLECTED' size='25' maxlength='35'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Check Amt Collected</b></font></td>
		<td>$<input type="text" name='CHECK_AMT_COLLECTED' size='10' maxlength='10'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Check Number</b></font></td>
		<td><input type="text" name='CHECK_NUM' size='10' maxlength='10'></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Money Off Coupon</b></font></td>
		<td>$<input type="text" name='DOLLARS_OFF_COUPON' size='40' maxlength='50'></td>
		</tr>

		<tr>
		<td align='right' valign='top'><font size='2' face='arial'><b>Notes</b></font></td>
		<td><textarea name='NOTES'></textarea></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Interested in volunteering?</b></font></td>
		<td>
			<select name='VOLUNTEER_INTEREST_IND'>
			<option value='N'>No</option>
			<option value='Y'>Yes</option>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Interested in being a captain (draft league only)?</b></font></td>
		<td>
			<select name='CAPTAIN_INTEREST_IND'>
			<option value='N'>No</option>
			<option value='Y'>Yes</option>
			</select>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Play Up/Down (draft league only)?</b></font></td>
		<td>
			<select name='PLAY_UPDOWN_IND'>
			<option value='N'>No</option>
			<option value='U'>Up</option>
			<option value='D'>Down</option>
			<option value='UD'>Up or Down</option>
			</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Registration'></font></td>
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
</body>
</html>
<%
'******************************************'

Sub closePage()
	rw("</body>")
	rw("</html>")
End Sub

'******************************************'

%>

