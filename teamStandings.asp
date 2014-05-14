<!-- #include virtual = "/incs/dbConnection.inc" -->
<!-- #include virtual="/incs/rw.asp" -->
<!-- BEGIN:teamStandings.asp -->
<html>
<!-- #include virtual="/incs/header.asp" -->
<%
	''rw("here")
	''Response.End
	EVENT_CD=Request("EVENT_CD")
	TEAM_ID=Request("TEAM_ID")
	DIVISION_ID=Request("DIVISION_ID")
	TEAM_NAME=Request("TEAM_NAME")

	WEB_PROC_NAME = "TEAM_RESULTS"
	WEB_PROC_VARS = "@EVENT_CD='" & EVENT_CD & "',@TEAM_ID='" & TEAM_ID & "', @DIVISION_ID='" & DIVISION_ID & "', @REPORT = 'N'"

	SQL = WEB_PROC_NAME & " " & WEB_PROC_VARS
	rw("<!-- SQL: " & SQL & "-->")

	''rw("SQL: " & SQL)
	''Response.End

	set rs = cn.Execute(SQL)

	if rs.EOF then
		rw("<table bgcolor='#FFFFFF' width='100%' height='300' align='center' border='0' cellspacing='1' cellpadding='4'>")

		rw("<tr valign='top'><td><font class='cfontError10'>No standings are available at this time. Please check back later.</font></td></tr>")

		rw("</table>")

	else
		rsData = rs.GetRows()
		rsCols = Ubound(rsData,1)
		rsRows = Ubound(rsData,2)

		dim headerArray()
		For x = 0 to rsCols
			redim preserve headerArray(x)
			headerArray(x) = rs(x).Name
		Next


		'Report column headers'
		rw("<table align='center'>")
		rw("<tr><td colspan='" & rsCols+1 & "' align='center' BGCOLOR='#000066'><font class='cfontWhite14'>Team Standings Detail - " & TEAM_NAME & "</font></td></tr>")

		rw("<tr>")

		For j = 0 to rsCols
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><nobr><font class='cfontWhite10'>" & headerArray(j) & "</font></nobr></th>")
		Next

		rw("</tr>")

		for i = 0 to rsRows

			rw("<tr>")

			ODD_ROW = NOT ODD_ROW

			If ODD_ROW Then
				BGCOLOR = "#FFFFFF"
			Else
				BGCOLOR = "#F0F0F0"
			End If

			For j = 0 to rsCols

			If j = 1 then
				 lalign="left"
			Else
				lalign="center"
			End If

			rw("<td align='" & lalign & "' bgcolor=""" & BGCOLOR & """><nobr><font class='cfont10'>" & rsData(j, i) & "</font></nobr></td>")

			Next

		next

		rw("</table>")

	end if
%>
</body>
</html>
<!-- END:teamStandings.asp -->
