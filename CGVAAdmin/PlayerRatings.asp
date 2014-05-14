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
<title>CGVA - Player Ratings Administration</title>

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

		for(i=0;i<MODIFY.ID.length;i++)
		{

			/*
			if(MODIFY.RATING_SCORE[i].value.length == 0)
			{
				alert("Please enter a rating score for this person.");
				MODIFY.RATING_SCORE[i].focus();
				return false;
			}

			if(MODIFY.RATING_PERSON_ID1[i].value == -1)
			{
				alert("Please enter at least one rater for this person.");
				MODIFY.RATING_PERSON_ID1[i].focus();
				return false;
			}

			if(MODIFY.EFF_DATE[i].value.length == 0)
			{
				alert("Please enter a date for this person's rating.");
				MODIFY.EFF_DATE[i].focus();
				return false;
			}
			*/
		}

	}
	else
	{

		/*
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
	<b><u>CGVA - Player Ratings Administration</u></b>
	</font>
	</div>

	<br />

	<%
		'call subroutines'
''		If Request("choice") = "" then
''			Call chooseOption()
''			Call closePage()
''		ElseIf Request("choice") = "modify" then
			Call modifyRecord()
			Call closePage()
''		ElseIf Request("choice") = "insert" then
''			Call insertRecord()
''			Call closePage()
''		End If


	'******************************************'

	Sub chooseOption()
	%>

		<form name="CHOICE" method="POST" action="PlayerRatings.asp">

		<table align='center' border='0' cellpadding='3'>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="insert" checked> Add New Player Rating
			</font>
		</td>
		</tr>

		<tr>
		<td>
			<font class='cfont10'>
			<input type="radio" name="choice" value="modify"> Modify Existing Player Ratings
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

		If Request("FILTER") = "" Then
			vFILTER = "A"
		Else
			vFILTER = Request("FILTER")
		End If

		sql = "SELECT 		a.PERSON_ID, a.FIRST_NAME, a.LAST_NAME, b.RATING_SCORE, "_
			& "				b.RATING_PERSON_ID1, b.RATING_PERSON_ID2, b.RATING_PERSON_ID3, b.EFF_DATE, b.NOTES "_
			& "FROM 		db_accessadmin.PERSON_TBL a "_
			& "LEFT JOIN 	RATINGS_TBL b on a.PERSON_ID = b.PERSON_ID "_
			& "WHERE 		a.LAST_NAME LIKE '" & Replace(vFILTER,"ALL","") & "%' "_
			& "AND	 		(b.EFF_DATE IS NULL OR b.EFF_DATE = (SELECT MAX(EFF_DATE) FROM RATINGS_TBL WHERE PERSON_ID = a.PERSON_ID)) "_
			& "ORDER BY 	UPPER(a.LAST_NAME), UPPER(a.FIRST_NAME), b.PERSON_ID, b.EFF_DATE"

		''rw(sql)
		''Response.End

		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsPlayerRatingData = rs.GetRows
			rsPlayerRatingRows = UBound(rsPlayerRatingData,2)
		else
			rw("Error:Missing Records.")
			Response.End
		end if


		SQL = "SELECT 		a.PERSON_ID, b.FIRST_NAME, b.LAST_NAME "_
			& "FROM 		RATER_TBL a LEFT JOIN db_accessadmin.PERSON_TBL b on a.PERSON_ID = b.PERSON_ID "_
			& "ORDER BY 	UPPER(b.LAST_NAME), UPPER(b.FIRST_NAME)"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

		if not rs.EOF then
			rsRaterData = rs.GetRows
			rsRaterRows = UBound(rsRaterData,2)
		else
			rsRaterRows = -1
		end if
	%>

		<form name='MODIFY' method='post' action='PlayerRatings_Submit.asp'>
		<div align='center'>
		<font class='cfont10'><b>
		Make any necessary changes to the player rating information.
		<br />
		Only the most recent rating information for players is listed below.
		<br />
		To update notes only, leave all other fields intact and add/modify any notes.
		<br />
		<font class='cfontError10'>Rating, Rater 1, and Effective Date are required.</font></b></font>
		<br /><br />

		<%
	''		If rs.EOF then
	''			rw("<font class='cfontError12'>* No data available to modify. *</font>")
	''			Response.End
	''		End If
		%>

			<select name='FILTER' onChange="changePage();">
			<option value='A'<%If vFILTER = "A" Then rw(" selected") End If%>>A</option>
			<option value='B'<%If vFILTER = "B" Then rw(" selected") End If%>>B</option>
			<option value='C'<%If vFILTER = "C" Then rw(" selected") End If%>>C</option>
			<option value='D'<%If vFILTER = "D" Then rw(" selected") End If%>>D</option>
			<option value='E'<%If vFILTER = "E" Then rw(" selected") End If%>>E</option>
			<option value='F'<%If vFILTER = "F" Then rw(" selected") End If%>>F</option>
			<option value='G'<%If vFILTER = "G" Then rw(" selected") End If%>>G</option>
			<option value='H'<%If vFILTER = "H" Then rw(" selected") End If%>>H</option>
			<option value='I'<%If vFILTER = "I" Then rw(" selected") End If%>>I</option>
			<option value='J'<%If vFILTER = "J" Then rw(" selected") End If%>>J</option>
			<option value='K'<%If vFILTER = "K" Then rw(" selected") End If%>>K</option>
			<option value='L'<%If vFILTER = "L" Then rw(" selected") End If%>>L</option>
			<option value='M'<%If vFILTER = "M" Then rw(" selected") End If%>>M</option>
			<option value='N'<%If vFILTER = "N" Then rw(" selected") End If%>>N</option>
			<option value='O'<%If vFILTER = "O" Then rw(" selected") End If%>>O</option>
			<option value='P'<%If vFILTER = "P" Then rw(" selected") End If%>>P</option>
			<option value='Q'<%If vFILTER = "Q" Then rw(" selected") End If%>>Q</option>
			<option value='R'<%If vFILTER = "R" Then rw(" selected") End If%>>R</option>
			<option value='S'<%If vFILTER = "S" Then rw(" selected") End If%>>S</option>
			<option value='T'<%If vFILTER = "T" Then rw(" selected") End If%>>T</option>
			<option value='U'<%If vFILTER = "U" Then rw(" selected") End If%>>U</option>
			<option value='V'<%If vFILTER = "V" Then rw(" selected") End If%>>V</option>
			<option value='W'<%If vFILTER = "W" Then rw(" selected") End If%>>W</option>
			<option value='X'<%If vFILTER = "X" Then rw(" selected") End If%>>X</option>
			<option value='Y'<%If vFILTER = "Y" Then rw(" selected") End If%>>Y</option>
			<option value='Z'<%If vFILTER = "Z" Then rw(" selected") End If%>>Z</option>
			<option value='ALL'<%If vFILTER = "ALL" Then rw(" selected") End If%>>All</option>
			</select>
		</div>


		<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

		<tr>
		<td colspan='7'><input type='submit' name="submitChoice" value='Modify Player Ratings(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Player</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Rating</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Rater 1</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Rater 2</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Rater 3</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Effective Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Notes</b></font></th>
		</tr>

		<%
			For i = 0 to rsPlayerRatingRows

				ODD_ROW = NOT ODD_ROW

				If ODD_ROW Then
					BGCOLOR = "#FFFFFF"
				Else
					BGCOLOR = "#F0F0F0"
				End If
		%>
				<tr bgcolor="<%=BGCOLOR%>">
				<td valign='top'>
					<input type='hidden' id='ID<%=rsPlayerRatingData(0,i)%>' name='ID' value='<%=rsPlayerRatingData(0,i)%>' />
					<font class='cfont10'><%=rsPlayerRatingData(2,i)%>, <%=rsPlayerRatingData(1,i)%></font>
				</td>
				<td valign='top' align='center'><input type='text' name='RATING_SCORE' ID='RATING_SCORE<%=rsPlayerRatingData(0,i)%>' value='<%=rsPlayerRatingData(3,i)%>' size='4' /></td>
				<td valign='top'>
					<select id="RATING_PERSON_ID1_<%=rsPlayerRatingData(0,i)%>" name="RATING_PERSON_ID1">
					<option value='-1'>-select-</option>
					<%
						for j = 0 to rsRaterRows
							rw("<option value='" & rsRaterData(0,j) & "'")

							If rsPlayerRatingData(4,i) = rsRaterData(0,j) then
								rw("selected")
							End If

							rw(">" & rsRaterData(2,j) & ", " & rsRaterData(1,j) & "(PID: " & rsRaterData(0,j) & ")</option>")
						next

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="RATING_PERSON_ID2_<%=rsPlayerRatingData(0,i)%>" name="RATING_PERSON_ID2">
					<option value='-1'>-select-</option>
					<%
						for j = 0 to rsRaterRows
							rw("<option value='" & rsRaterData(0,j) & "'")

							If rsPlayerRatingData(5,i) = rsRaterData(0,j) then
								rw("selected")
							End If

							rw(">" & rsRaterData(2,j) & ", " & rsRaterData(1,j)  & "(PID: " & rsRaterData(0,j) & ")</option>")
						next

					%>

					</select>
				</td>
				<td valign='top'>
					<select id="RATING_PERSON_ID3_<%=rsPlayerRatingData(0,i)%>" name="RATING_PERSON_ID3">
					<option value='-1'>-select-</option>
					<%
						for j = 0 to rsRaterRows
							rw("<option value='" & rsRaterData(0,j) & "'")

							If rsPlayerRatingData(6,i) = rsRaterData(0,j) then
								rw("selected")
							End If

							rw(">" & rsRaterData(2,j) & ", " & rsRaterData(1,j)  & "(PID: " & rsRaterData(0,j) & ")</option>")
						next

					%>

					</select>
				</td>
				<td valign='top' align='center'><input type='text' name='EFF_DATE' ID='EFF_DATE<%=rsPlayerRatingData(0,i)%>' value='<%If IsNull(rsPlayerRatingData(7,i)) Then rw("") Else rw(FormatDateTime(rsPlayerRatingData(7,i),2)) End If%>' size='10' /></td>
				<td valign='top' align='center'><textarea id="ID<%=rsPlayerRatingData(0,i)%>" name="NOTES"><%=rsPlayerRatingData(8,i)%></textarea></td>
				</tr>

		<%
				Next
		%>

		<tr>
		<td colspan='7'><input type='submit' name="submitChoice" value='Modify Player Ratings(s)' onClick="return validateModify();"></td>
		</tr>

		</table>

		</form>

		<br /><br />

		<script language='Javascript'>
		<!--
			MODIFY.FILTER.focus();

			function changePage()
			{
				vFILTER = MODIFY.FILTER.options[MODIFY.FILTER.selectedIndex].value;
				//alert(choice);
				MODIFY.submit();
			}
		//-->
		</script>


	<%
		''call closeRSCNConnection()
	End Sub

	'******************************************'

	Sub insertRecord()
	''not used

	%>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the new rater information.
		</b></font>
		</div>

		<br /><br />

		<form name='INSERT' method='post' action='PlayerRatings_submit.asp' onSubmit="return validateInsert();">

		<table width='500' bgcolor='#F0F0F0' cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'><b>Rater</b></font></td>
		<td>
			<select name='RATER'>
			<%
				For i = 0 to rsPlayerRatingRows
					rw("<option value='" & rsPlayerRatingData(0,i) & "'>" & rsPlayerRatingData(2,i) & ", " & rsPlayerRatingData(1,i) & "(PID: " & rsPlayerRatingData(0,i) & ")</option>")
				Next
			%>
			</select>
		</td>
		</tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><input type='submit' name="submitChoice" value='Add Player Rating'></font></td>
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


