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
<title>CGVA - Player Ref Certification Administration</title>

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
			if(MODIFY.REF_SCORE[i].value.length == 0)
			{
				alert("Please enter a ref certification for this person.");
				MODIFY.REF_SCORE[i].focus();
				return false;
			}

			if(MODIFY.EFF_DATE[i].value.length == 0)
			{
				alert("Please enter a date for this person's ref certification.");
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
	var CertType = eval("MODIFY.CertType" + rowValue);

	if(idValue.checked)
	{
		CertType.disabled = false;
	}
	else
	{
		CertType.disabled = true;
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
	<b><u>CGVA - Player Ref Certification Administration</u></b>
	</font>
	</div>

	<br />

	<%
		'call subroutines'
		Call modifyRecord()
		Call closePage()

	'******************************************'

	Sub modifyRecord()

		If Request("FILTER") = "" Then
			vFILTER = "A"
		Else
			vFILTER = Request("FILTER")
		End If

		sql = "SELECT 		a.PERSON_ID, a.FIRST_NAME, a.LAST_NAME, b.CERT_ID, "_
			& "				b.EFF_DATE, b.END_DATE,b.NOTES "_
			& "FROM 		db_accessadmin.PERSON_TBL a "_
			& "LEFT JOIN 	REFEREE_CERTIFICATION_TBL b on a.PERSON_ID = b.PERSON_ID "_
			& "LEFT JOIN 	CERTIFICATION_TYPE_TBL c on b.CERT_ID = c.CERT_ID "_
			& "WHERE 		a.LAST_NAME LIKE '" & Replace(vFILTER,"ALL","") & "%' "_
			& "AND	 		(b.END_DATE IS NULL OR b.END_DATE = (SELECT MAX(END_DATE) FROM REFEREE_CERTIFICATION_TBL WHERE PERSON_ID = a.PERSON_ID)) "_
			& "ORDER BY 	UPPER(a.LAST_NAME), UPPER(a.FIRST_NAME), b.PERSON_ID, b.END_DATE"

		''rw(sql)
		''Response.End

		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsPlayerRefData = rs.GetRows
			rsPlayerRefRows = UBound(rsPlayerRefData,2)
		else
			rw("Error:Missing Records.")
			Response.End
		end if


		SQL = "SELECT 		CERT_ID, CERT_DESC "_
			& "FROM 		CERTIFICATION_TYPE_TBL "_
			& "ORDER BY 	UPPER(CERT_DESC)"
		rw("<!-- SQL: " & SQL & " -->")
		Set rs = cn.Execute(SQL)

		if not rs.EOF then
			rsCertTypeData = rs.GetRows
			rsCertTypeRows = UBound(rsCertTypeData,2)
		else
			rsCertTypeRows = -1
		end if
	%>

		<form name='MODIFY' method='post' action='PlayerRefCertification_Submit.asp'>
		<div align='center'>
		<font class='cfont10'><b>
		Make any necessary changes to the player ref certification information.
		<br />
		Only the most recent ref certification information for players is listed below.
		<br />
		To update notes only, leave all other fields intact and add/modify any notes.
		<br />
		<font class='cfontError10'>Certification Type and Effective Date are required.</font></b></font>
		<br /><br />

		<%
			If Session("Err") <> "" then
				rw("<font class='cfontError12'>" & Session("Err") & "</font>")
				Session("Err") = ""
				rw("<br />")
			End If
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
		<td colspan='4'><input type='submit' name="submitChoice" value='Modify Player Ref Certification(s)' onClick="return validateModify();"></td>
		</tr>

		<tr bgcolor='#000066'>
		<th valign='bottom'><font class='cfontWhite10'><b>Player</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Certification</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Effective Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>End Date</b></font></th>
		<th valign='bottom'><font class='cfontWhite10'><b>Notes</b></font></th>
		</tr>

		<%
			For i = 0 to rsPlayerRefRows

				ODD_ROW = NOT ODD_ROW

				If ODD_ROW Then
					BGCOLOR = "#FFFFFF"
				Else
					BGCOLOR = "#F0F0F0"
				End If
		%>
				<tr bgcolor="<%=BGCOLOR%>">
				<td valign='top'>
					<input type='hidden' id='ID<%=rsPlayerRefData(0,i)%>' name='ID' value='<%=rsPlayerRefData(0,i)%>' />
					<font class='cfont10'><%=rsPlayerRefData(2,i)%>, <%=rsPlayerRefData(1,i)%> (PID: <%=rsPlayerRefData(0,i)%>)</font>
				</td>
				<td valign='top'>
					<select id="CERT_ID<%=rsPlayerRefData(0,i)%>" name="CERT_ID">
					<option value='-1'>-select-</option>
					<%
						for j = 0 to rsCertTypeRows
							rw("<option value='" & rsCertTypeData(0,j) & "'")

							If rsPlayerRefData(3,i) = rsCertTypeData(0,j) then
								rw("selected")
							End If

							rw(">" & rsCertTypeData(1,j) & "</option>")
						next

					%>

					</select>
				</td>
				<td valign='top' align='center'><input type='text' name='EFF_DATE' ID='EFF_DATE<%=rsPlayerRefData(0,i)%>' value='<%If IsNull(rsPlayerRefData(4,i)) Then rw("") Else rw(FormatDateTime(rsPlayerRefData(4,i),2)) End If%>' size='10' /></td>
				<td valign='top' align='center'><input type='text' name='END_DATE' ID='END_DATE<%=rsPlayerRefData(0,i)%>' value='<%If IsNull(rsPlayerRefData(5,i)) Then rw("") Else rw(FormatDateTime(rsPlayerRefData(5,i),2)) End If%>' size='10' /></td>
				<td valign='top' align='center'><textarea id="ID<%=rsPlayerRefData(0,i)%>" name="NOTES"><%=rsPlayerRefData(6,i)%></textarea></td>
				</tr>

		<%
				Next
		%>

		<tr>
		<td colspan='4'><input type='submit' name="submitChoice" value='Modify Player Ref Certification(s)' onClick="return validateModify();"></td>
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


