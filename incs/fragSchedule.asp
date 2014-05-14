<!-- BEGIN:fragSchedule.asp -->
<%
	WEB_PROC_NAME = "MATCH_SCHEDULE"
	''WEB_PROC_VARS = "@EVENT_CD='" & EVENT_CD & "',@DIVISION_ID='" & DIVISION_ID & "',@SORT='" & Request("SORT") & "'"
	WEB_PROC_VARS = "@EVENT_CD='" & EVENT_CD & "',@DIVISION_ID='" & DIVISION_ID & "',@SORT='" & Request("SORT") & "', @SCHEDULE_CHOICE='" & Session("SCHEDULE_CHOICE") & "'"

	SQL = WEB_PROC_NAME & " " & WEB_PROC_VARS
	rw("<!-- SQL: " & SQL & "-->")

	''rw("SQL: " & SQL)
	''Response.End

	set rs = cn.Execute(SQL)

	if rs.EOF then
		rw("<table bgcolor='#FFFFFF' width='100%' height='300' align='center' border='0' cellspacing='1' cellpadding='4'>")

		rw("<tr valign='top'><td><font class='cfontError10'>No schedule is available at this time. Please check back later.</font></td></tr>")

		rw("</table>")

	else
		rsData = rs.GetRows()
		rsCols = Ubound(rsData,1)
		rsRows = Ubound(rsData,2)

		rw("<table width='100%' align='center' border='0' cellspacing='1' cellpadding='4' bgcolor='#9999FF'>")

		'Report column headers'
			rw("<tr>")

			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'>WEEK</font></th>")
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'><a class='menuWhite' href='index.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'>TIME</a></font></th>")
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'><a class='menuWhite' href='index.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=COURT'>COURT</a></font></th>")
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'>PLAY</font></th>")
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'>PLAY</font></th>")
			rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'>REF</font></th>")
			rw("</tr>")


		'Report column headers'

		WEEK_NUM = ""
		for i = 0 to rsRows

			rw("<tr>")

			If WEEK_NUM <> rsData(0,i) Then
			    WEEK_NUM = rsData(0,i)
			    ODD_ROW = NOT ODD_ROW
			    SHOW_WEEK = TRUE
            Else
                SHOW_WEEK = FALSE
            End If

			If ODD_ROW Then
				BGCOLOR = "#FFFFFF"
			Else
				BGCOLOR = "#CFCFCF"
			End If

			For j = 0 to rsCols

				lalign="center"

				If j = 0 and SHOW_WEEK = FALSE then
					rw("<td align='" & lalign & "' bgcolor=""" & BGCOLOR & """><nobr><font class='cfont10'>&nbsp;</font></nobr></td>")
				Else
					rw("<td align='" & lalign & "' bgcolor=""" & BGCOLOR & """><nobr><font class='cfont10'>" & rsData(j, i) & "</font></nobr></td>")
				End If

			Next

    		rw("</tr>")

		next

		rw("</table>")

	end if
%>



	<table width='100%'>


</table>
<!-- END:fragSchedule.asp -->



