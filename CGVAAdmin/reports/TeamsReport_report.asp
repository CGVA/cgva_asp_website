<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	ON ERROR RESUME NEXT
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	''Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	Dim TYear, Month
	Dim DEBUG
	DIM START_QUERY_TIME, STOP_QUERY_TIME, START_PAGE_TIME	'For troubleshooting purposes'
	DIM EDATE, SDATE, EXCEL									'Values passed from previous page'
	Dim numcols, numrows, alldata, header,allJdata,allJrows
	Dim TYPEOFREPORT, vEVENT
	Dim STOP_DATE, START_DATE, STOP_TIME, START_TIME
	Dim SQL
	Dim WEB_PROC_NAME						'The proc we are going to run'
	Dim WEB_PROC_VARS                       'The variables required by the web proc'

	DEBUG = false  'set this to TRUE to debug'

	START_PAGE_TIME = Now

''	If Session("userID") = "" then
''		Session("Err") = "Your session has timed out. Please log in again."
''		Response.Redirect("/logout.asp")
''	End If



	'****************'
	'*    MAIN      *'
	'****************'

	Session("EVENT") = Request("EVENT")
	vEVENT = Request("EVENT")
	Session("DISPLAY") = Request("DISPLAY")

	If Request("DISPLAY") = "EXCEL" Then
		EXCEL = "ON"
	End If

	Session("showname") = "Team Information"
	WEB_PROC_NAME = "TEAM"
	WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"

	''rw("wpv:" & WEB_PROC_VARS)
	''Response.End

	SQL = WEB_PROC_NAME & " " & WEB_PROC_VARS
	rw("<!-- SQL: " & SQL & "-->")

	''rw("SQL: " & SQL)
	''Response.End

	set rs = cn.Execute(SQL)

	If EXCEL = "ON" Then
		Response.ContentType = "application/vnd.ms-excel"
	Else
%>
		<!-- #include virtual="/incs/fragHeader.asp" -->
		<title>CGVA - Reports</title>
		</head>

		<!-- #include virtual="/incs/rw.asp" -->
		<!-- #include virtual="/incs/header.asp" -->
		<!-- #include virtual="/incs/fragHeaderReportGraphics.asp" -->


		<tr>
		<td valign='top' width='100%'>

			<div align='center'>
			<font class='cfont14'>
			<b><u>CGVA - Reports</u></b>
			</font>
			<br /><br />
<%
	End If


	If rs.EOF Then
		rw("<div align='center'><font class='cfontError10'><b>No data was returned for this report.</b></font></div>")

		'Page footer'
		If EXCEL <> "ON" Then
			rw("<p align='center'>")
			rw("<font class='cfont10'><a class='menuGray' href='CGVAReports.asp'>Select Another Report</a>&nbsp;</font>")
			rw("</p>")
		End If

		closeRSCNConnection()
		Response.End
	Else

		rsData = rs.GetRows()
		rsCols = Ubound(rsData,1)
		rsRows = Ubound(rsData,2)

		If EXCEL <> "ON" Then
			call WebReport()
		Else
			call ExcelReport()
		End If

	End If

	'888888888888888888888888888888888888888888888888888888888888888888'

	Sub WebReport()
		rw("<table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='#ffffff'>")

		rw("<tr>")
		rw("<td align='center'>")

			rw("<div align='center'>")
			rw("<font class='cfont12'><b><u>" & Session("showname") & "</u></b></font>")
			rw("</div>")

			rw("<br />")

			rw("<table width='98%' align='center' border='0' cellspacing='1' cellpadding='4' bgcolor='#FFFFFF'>")

			rw("<tr>")
			rw("<td valign='top'>")

				rw("<table align='center' border='0' cellspacing='1' cellpadding='4' bgcolor='#9999FF'>")

				'Report column headers'
				rw("<tr>")

				For j = 0 to rsCols
					rw("<th align='center' valign='bottom' BGCOLOR='#000066'><font class='cfontWhite10'>" & rs(j).Name & "</font></th>")
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
						rw("<td bgcolor=""" & BGCOLOR & """><font class='cfont10'>" & rsData(j, i) & "</font></td>")
					Next

					rw("</tr>")

				next

				rw("</table>")

			rw("</td>")
			rw("</tr>")

			rw("</table>")

			rw("<br /><br />")

			'Page footer'
			If EXCEL <> "ON" Then
				rw("<p align='center'>")
				rw("<font class='cfont10'><a class='menuGray' href='CGVAReports.asp'>Select Another Report</a>&nbsp;</font>")
				rw("</p>")
			End If

		rw("</td>")
		rw("</tr>")

		rw("</table>")

rw("</td>")
rw("</tr>")

rw("<tr height='20'><td></td></tr>")
%>
<!--#include virtual = "/incs/fragContact.asp"-->
<%
rw("</table>")

		closeRSCNConnection()
	End Sub

	'888888888888888888888888888888888888888888888888888888888888888888'

	Sub ExcelReport()
		rw("<table width='100%' border='0' cellpadding='0' cellspacing='0' bgcolor='#ffffff'>")

		rw("<tr>")
		rw("<td align='center'>")

			rw("<div align='center'>")
			rw("<font class='cfont14'><b><u>" & Session("showname") & "as of " & Now() & "</u></b></font>")
			rw("</div>")

			rw("</td></tr></table>")
			rw("<br />")

			rw("<table align='center' border='1' cellspacing='1' cellpadding='4'>")

			'Report column headers'
			rw("<tr>")

			For j = 0 to rsCols
				rw("<th align='center' valign='bottom'><font face='arial' size='2'>" & rs(j).Name & "</font></th>")
			Next

			rw("</tr>")

			for i = 0 to rsRows

				rw("<tr>")

				For j = 0 to rsCols
					rw("<td><font face='arial' size='2'>" & rsData(j, i) & "</font></td>")
				Next

				rw("</tr>")

			next

			rw("</table>")

		closeRSCNConnection()
	End Sub


	'888888888888888888888888888888888888888888888888888888888888888888'

%>


<br />
<br />

</body>
</html>

