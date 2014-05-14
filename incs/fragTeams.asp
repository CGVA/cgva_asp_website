<!-- BEGIN:fragTeams.asp -->
<%
	SQL = "TEAM_WEB @EVENT_CD='" & EVENT_CD & "',@DIVISION_ID='" & DIVISION_ID & "'"
	set rs = cn.Execute(SQL)

	if not rs.EOF then
		rsTeamData = rs.GetRows()
		rsTeamRows = UBound(rsTeamData,2)
	else
		rsTeamRows = -1
	end if


%>
<table width='100%' cellspacing='1' cellpadding='2' bgcolor='#9999FF'>

<tr bgcolor='#000066'>
<th><font class='cfontWhite10'>Team Code</font></th>
<th><font class='cfontWhite10'>Team Name</font></th>
<th><font class='cfontWhite10'>Last Name</font></th>
<th><font class='cfontWhite10'>First Name</font></th>
<th><font class='cfontWhite10'>Captain</font></th>
<th><font class='cfontWhite10'>Certified Ref</font></th>
</tr>

<%
	If rsTeamRows > -1 Then

		TEMPTEAM = ""

		For i = 0 to rsTeamRows
			rw("<tr bgcolor='#FFFFFF'>")

			If rsTeamData(0,i) <> TEMPTEAM Then
				ODD_ROW = NOT ODD_ROW

				If ODD_ROW Then
					BGCOLOR = "#FFFFFF"
				Else
					BGCOLOR = "#CFCFCF"
				End If
				TEMPTEAM = rsTeamData(0,i)
				rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(0,i) & "&nbsp;</font></td>")
				rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(1,i) & "&nbsp;</font></td>")
			Else
				rw("<td bgcolor='#9999FF' align='center' colspan='2'><font class='cfont10'>&nbsp;</font></td>")
			End If

			rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(4,i) & "&nbsp;</font></td>")
			rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(5,i) & "&nbsp;</font></td>")
			rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(2,i) & "&nbsp;</font></td>")
			rw("<td BGCOLOR='" & BGCOLOR & "' align='center'><font class='cfont10'>" & rsTeamData(3,i) & "&nbsp;</font></td>")
			rw("</tr>")
		Next

	Else
			rw("<tr height='200' bgcolor='#FFFFFF'>")
			rw("<td valign='top' colspan='6' align='center'><font class='cfontError10'>No team rosters are available at this time.</font></td>")
			rw("</tr>")
	End If

%>
</table>
<!-- END:fragTeams.asp -->
