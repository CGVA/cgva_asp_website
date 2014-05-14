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
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	''check for submit
	if request("submit") <> "" then

		ID = Request("ID")
		IDarray = split(ID,", ")

		WEEK_NUM_DISPLAY_IND = FixEmptyCell(Request("WEEK_NUM_DISPLAY_IND"))
		WEEK_NUM_DISPLAY_INDarray = split(WEEK_NUM_DISPLAY_IND,", ")

		For i = 0 to UBound(IDarray)
			sql = "UPDATE EVENT_TBL SET "_
				& "WEEK_NUM_DISPLAY_IND = '" & WEEK_NUM_DISPLAY_INDarray(i) & "' "_
				& "WHERE EVENT_CD = '" & IDarray(i) & "'"

			''rw(IDarray(i) & "<br />")
			''rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)

		Next

		Session("Err") = "Y"
		''Response.End

	end if


	SQL = "SELECT e.EVENT_CD, "_
		& "l.LOCATION_DESC, "_
		& "e.EVENT_SHORT_DESC + '/' + e.EVENT_LONG_DESC as 'EVENT', "_
		& "e.WEEK_NUM_DISPLAY_IND "_
		& "FROM 		EVENT_TBL e "_
		& "LEFT JOIN LOCATION_TBL l "_
		& "ON e.LOCATION_CD = l.LOCATION_CD "_
		& "WHERE e.ACTIVE_EVENT_IND = 'Y' "_
		& "ORDER BY EVENT"
	rw("<!-- SQL: " & SQL & " -->")
	Set rs = cn.Execute(SQL)

	if not rs.EOF then
		rsEventData = rs.GetRows
		rsEventRows = UBound(rsEventData,2)
	else
		rsEventRows = -1
	end if

%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Active Period Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Active Period Administration</u></b>
	</font>
	</div>

	<br />

	<div align='center'>
	<font class='cfont10'><b>Update the active period indicator for a listed event.</b></font>
	<br />
	<font class='cfont10'><b>cgva.org will display results up to and including the period selected for the event.</b></font>
	<br /><br />

	<%
		If Session("Err") <> "" then
			rw("<font class='cfontSuccess10'><b>The active periods were modified successfully.</b></font>")
			Session("Err") = ""
		End If
	%>

	</div>


	<form method='post' action='ActivePeriod.asp'>


	<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>

	<tr>
	<td colspan='3'><input type='submit' name="submit" value='Modify Active Period(s)'></td>
	</tr>

	<tr bgcolor='#000066'>
	<th valign='bottom'><font class='cfontWhite10'><b>Location</b></font></th>
	<th valign='bottom'><font class='cfontWhite10'><b>Event</b></font></th>
	<th valign='bottom'><font class='cfontWhite10'><b>Approved Period (for website display)</b></font></th>
	</tr>

	<%
		For i = 0 to rsEventRows

			ODD_ROW = NOT ODD_ROW

			If ODD_ROW Then
				BGCOLOR = "#FFFFFF"
			Else
				BGCOLOR = "#F0F0F0"
			End If

		''get possible period values for each active event as it is displayed
		''	& "w.EVENT_CD + ' - ' + CONVERT(varchar,w.WEEK_NUM) AS 'EVENT/WEEK#' "_
		sql = "SELECT w.WEEK_ID, "_
			& "w.WEEK_NUM AS 'PERIOD #' "_
			& "FROM WEEK_TBL w "_
			& "LEFT JOIN EVENT_TBL e "_
			& "ON w.EVENT_CD = e.EVENT_CD "_
			& "WHERE e.ACTIVE_EVENT_IND = 'Y' "_
			& "AND w.EVENT_CD = '" & rsEventData(0,i) & "' "_
			& "ORDER BY 'PERIOD #'"

		set rs = cn.Execute(sql)

		if not rs.EOF then
			rsWNDData = rs.GetRows
			rsWNDRows = UBound(rsWNDData,2)
		else
			rsWNDRows = -1
		end if
	%>

			<tr bgcolor="<%=BGCOLOR%>">
			<td><font class='cfont10'><%=rsEventData(1,i)%></font></td>
			<td><font class='cfont10'><%=rsEventData(2,i)%></font></td>
			<td>
				<select id="WEEK_NUM_DISPLAY_IND<%=rsEventData(0,i)%>" name="WEEK_NUM_DISPLAY_IND">
				<option value='0'>-none-</option>
				<%
					for j = 0 to rsWNDRows
						rw("<option value='" & rsWNDData(0,j) & "'")

						If rsEventData(3,i) = rsWNDData(0,j) then
							rw("selected")
						End If

						rw(">" & rsWNDData(1,j) & "</option>")
					next

				%>

				</select>
				<input type="hidden" id="ID<%=rsEventData(0,i)%>" name="ID" value="<%=rsEventData(0,i)%>">
			</td>
			</tr>

	<%
		Next
	%>

	<tr>
	<td colspan='3'><input type='submit' name="submit" value='Modify Active Period(s)'></td>
	</tr>

	</table>

	</form>

	<br /><br />


<%
	call closeRSCNConnection()
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
Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function

%>


