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

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"BOD") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If



	'****************'
	'*    MAIN      *'
	'****************'

	Session("DB") = Request("DB")
	Session("TYPEOFREPORT") = Request("TYPEOFREPORT")
	TYPEOFREPORT = Request("TYPEOFREPORT")
	Session("EVENT") = Request("EVENT")
	vEVENT = Request("EVENT")
	Session("DISPLAY") = Request("DISPLAY")

	If Request("DISPLAY") = "EXCEL" Then
		EXCEL = "ON"
	End If

	Session("DATETYPE") = Request("DATETYPE")


	If Request("DATETYPE") = "monthly" Then
		Session("SDATE") = Request("SDATE")
		Session("EDATE") = Request("EDATE")
		Session("MDATE") = Request("MDATE")

		SDATE = Request("MDATE")
		EDATE = DateAdd("m",1,Request("MDATE")) - 1

	ElseIf Request("DATETYPE") = "daily" Then
		Session("SDATE") = Request("SDATE")
		SDATE = Request("SDATE")
		Session("EDATE") = Request("EDATE")
		EDATE = Request("EDATE")
		Session("STIME") = Request("STIME")
		START_TIME = Request("STIME")
		Session("ETIME") = Request("ETIME")
		STOP_TIME = Request("ETIME")
	Else
		Session("SDATE") = Request("SDATE")
		SDATE = Request("SDATE")
		Session("EDATE") = Request("EDATE")
		EDATE = Request("EDATE")
	End If

	STOP_DATE = EDATE & " 23:59:59"
	START_DATE = SDATE

	If Len(SDATE) < 10 Then

		If Request("DATETYPE") = "monthly" Then

			TYear = Right(SDATE, 4)
			Month = Left(SDATE,3)

			Select Case Month
				Case "Jan"
					SDATE = "01/01/" & TYear
				Case "Feb"
					SDATE = "02/01/" & TYear
				Case "Mar"
					SDATE = "03/01/" & TYear
				Case "Apr"
					SDATE = "04/01/" & TYear
				Case "May"
					SDATE = "05/01/" & TYear
				Case "Jun"
					SDATE = "06/01/" & TYear
				Case "Jul"
					SDATE = "07/01/" & TYear
				Case "Aug"
					SDATE = "08/01/" & TYear
				Case "Sep"
					SDATE = "09/01/" & TYear
				Case "Oct"
					SDATE = "10/01/" & TYear
				Case "Nov"
					SDATE = "11/01/" & TYear
				Case "Dec"
					SDATE = "12/01/" & TYear
			End Select

			START_DATE = Trim(SDATE)
			STOP_DATE = Trim(EDATE) & " 23:59:59"

		End If

	End If


	'report header'
	Select Case Session("TYPEOFREPORT")

	Case "REGISTRATION"
		Session("showname") = "Registered Players"
		WEB_PROC_NAME = "REGISTRATION"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "', @START_DATE = '" & START_DATE & "', @STOP_DATE = '" & STOP_DATE & "'"
	Case "CONTACTS"
		Session("showname") = "Contact Information"
		WEB_PROC_NAME = "CONTACTS"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case "TEAM"
		Session("showname") = "Team Information"
		WEB_PROC_NAME = "TEAM"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case "RATINGS"
		Session("showname") = "Player Ratings"
		WEB_PROC_NAME = "RATINGS"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case "FASTLANE"
		Session("showname") = "Fast Lane Personnel"
		WEB_PROC_NAME = "FASTLANE"
		''WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case "RESULTS"
		Session("showname") = "Event Results"
		WEB_PROC_NAME = "RESULTS"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case "SCHEDULE"
		Session("showname") = "Event Schedule"
		WEB_PROC_NAME = "SCHEDULE"
		WEB_PROC_VARS = "@EVENT_CD='" & vEVENT & "'"
	Case Else
		Session("showname") = "Error: Type of Report not accounted for."
	End Select


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
		<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->


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
			rw("<font class='cfont10'><a class='menuGray' href='reports.asp'>Select Another Report</a>&nbsp;</font>")
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

				Select Case Session("TYPEOFREPORT")

				Case "REGISTRATION","CONTACTS","RATINGS","FASTLANE","TEAM","SCHEDULE"
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

				Case "RESULTS"
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

						DIV_ID = ""
						For j = 0 to rsCols

							''ranking teams
							If j = 0 then

							ElseIf j = 4 then
								rsData(j, i) = FormatNumber(rsData(j, i),3)
							End If

							rw("<td bgcolor=""" & BGCOLOR & """><font class='cfont10'>" & rsData(j, i) & "</font></td>")
						Next

						rw("</tr>")

					next

					rw("</table>")

				End Select

			rw("</td>")
			rw("</tr>")

			rw("</table>")

			rw("<br /><br />")

			'Page footer'
			If EXCEL <> "ON" Then
				rw("<p align='center'>")
				rw("<font class='cfont10'><a class='menuGray' href='reports.asp'>Select Another Report</a>&nbsp;</font>")
				rw("</p>")
			End If

		rw("</td>")
		rw("</tr>")

		rw("</table>")

rw("</td>")
rw("</tr>")

rw("<tr height='20'><td></td></tr>")
%>
<!--#include virtual="/incs/fragContact.asp"-->
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
			rw("<font class='cfont14'><b><u>" & Session("showname") & " as of " & Now() & "</u></b></font>")
			rw("</div>")

			rw("</td></tr></table>")
			rw("<br />")

		Select Case Session("TYPEOFREPORT")

		Case "REGISTRATION","CONTACTS","RATINGS", "FASTLANE","TEAM","SCHEDULE"
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

		Case "RESULTS"
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
					If j = 4 then
						rsData(j, i) = FormatNumber(rsData(j, i),3)
					End If
					rw("<td><font face='arial' size='2'>" & rsData(j, i) & "</font></td>")
				Next

				rw("</tr>")

			next

			rw("</table>")

		End Select

		closeRSCNConnection()
	End Sub


	'888888888888888888888888888888888888888888888888888888888888888888'

%>


<br />
<br />

</body>
</html>

