<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

''	If Session("userID") = "" then
''		Session("Err") = "Your session has timed out. Please log in again."
''		Response.Redirect("/logout.asp")
''	End If


	sql = 	"SELECT EVENT_CD, " & _
			"EVENT_SHORT_DESC " & _
			"FROM EVENT_TBL " & _
			"ORDER BY EVENT_SHORT_DESC"

	set rs = cn.Execute(sql)

	if rs.EOF then
		rw("Error: No events are available from which to select.")
		Response.End
	else
		rsEventData = rs.GetRows
		rsEventRows = UBound(rsEventData,2)
	end if

%>


<!-- #include virtual = "/incs/fragHeader.asp" -->
<title>CGVA - Reports</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual = "/incs/fragHeaderReportGraphics.asp" -->


<tr>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Teams Report</u></b>
	</font>
	</div>

	<br />
	<form method="POST" action="TeamsReport_report.asp" id="rptVariables" name="rptVariables" onSubmit="return checkVariables();">

	<!-- BEGIN MAIN BODY (2 panes - left & right) -->
	<table align='center' border="0" cellpadding="0" cellspacing="0">

	<tr>
	<!-- BEGIN LEFT BODY CONTENT  -->
	<td align="left" valign="top">

		<!-- BEGIN reportVariables -->
		<table width="100%" border="0" cellpadding="0" cellspacing="0">

		<tr>
		<td height="8"><img src="../../images/spacer.gif" width="1" height="8"></td>
		</tr>

		<tr>
		<td align="center" valign="middle"><font class='cfont10'><b>&nbsp;Select the event you wish to view and click submit.</b></font></td>
		</tr>

		</table>


		<table width="100%" cellpadding="0" cellspacing="0" border="0">

		<tr>
		<td colspan='3' align="left" valign="middle"><img src="../../images/spacer.gif" WIDTH="30" HEIGHT="20"></td>
		</tr>

		<tr>
		<td align="left" valign="top">

			<!-- REPORT VARIABLE DISPLAY FOR THE OLS SYSTEM -->
			<table width="100%" cellpadding="3" cellspacing="0" border="0">

			<!-- REPORT VARIABLE - REQUIRED FOR ALL REPORTS -->
			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../../images/rpt_ReportOptions.gif" width="235" height="15"></td>
			</tr>

			<tr>
			<td width='1%'><font class='cfont10'>Event:</font></td>
			<td align='left'>
				<select name="EVENT" size="1">
				<% For i = 0 to rsEventRows%>
					<option value="<%=rsEventData(0,i)%>"<%If Session("EVENT") = rsEventData(0,i) Then rw(" SELECTED") %>><%=rsEventData(1,i)%></option>
				<%Next%>
				</select>
			</td>
			</tr>

			<!-- END REPORT VARIABLE -->

			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../../images/spacer.gif" HEIGHT="20"></td>
			</tr>

			<!-- SCROLLING, PRINTER FRIENDLY OR EXCEL DISPLAY -->
			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../../images/rpt_DisplayOptions.gif" width="235" height="15"></td>
			</tr>

			<tr>
			<td colspan='2' align="left" valign="top">
				<input type="radio" name="DISPLAY" value="PRINT" <%If Session("DISPLAY")="" or Session("DISPLAY")="PRINT" then%> checked<%End If%>><font class='cfont10'>Screen</font>&nbsp;
				<input type="radio" name="DISPLAY" value="EXCEL" <%If Session("DISPLAY")="EXCEL" then%> checked<%End If%>><font class='cfont10'>Excel</font>
			</td>
			</tr>
			<!-- END SCROLLING, PRINTER FRIENDLY OR EXCEL DISPLAY -->


			</table>

		</td>
		<!-- END REPORT VARIABLE DISPLAY FOR THE OLS SYSTEM -->

		<!-- SPACER -->
		<td width="5"><img src="../../images/spacer.gif" width="5" height="1"></td>
		<!-- END SPACER -->

		<!-- SCROLLING, PRINTER FRIENDLY OR EXCEL DISPLAY -->
		<td align="left" valign="top">
			&nbsp;
		</td>
		</tr>

		<!-- SPACER -->
		<tr>
		<td colspan="3"><img src="../../images/spacer.gif" height="10"></td>
		</tr>
		<!-- END SPACER -->

		</table>


		<!-- VARIABLE FOOTER (Error Control and Buttons) -->
		<table width="100%" cellpadding="0" cellspacing="0" border="0">

		<!-- SESSION ERROR DISPLAY - HIDDEN UNLESS USER ENCOUNTERS AN ERROR -->
		<tr>
		<td colspan="3" ALIGN="RIGHT" VALIGN="TOP" NOWRAP>
			<p align="center"><% = Session("Err") %><% Session("Err") = "" %>&nbsp;</p>
		</td>
		</tr>
		<!-- END SESSION ERROR DISPLAY -->

		<!-- SUBMIT/RESET BUTTONS -->
		<tr>
		<td colspan="3" height="10"><img src="../../images/spacer.gif" width="1" height="10"></td>
		</tr>

		<tr>
		<td width="25" height="1"><img src="../../images/spacer.gif" width="25" height="1"></td>
		<td width="450" height="1" bgcolor="black"><img src="../../images/spacer.gif" width="1" height="1"></td>
		<td width="25" height="1"><img src="../../images/spacer.gif" width="25" height="1"></td>
		</tr>

		<tr>
		<td colspan="3" height="10"><img src="../../images/spacer.gif" width="1" height="10"></td>
		</tr>

		<tr>
		<td colspan="3" align="center">
			<input type="submit" value="Submit" name="SUBMIT">&nbsp;
			<input type="reset" value="Reset" name="B2">
		</td>
		</tr>
		<!-- END SUBMIT/RESET BUTTONS -->

		</table>



	</td>
	</tr>

	</table>

</td>
</tr>

<tr height='20'><td></td></tr>
<!--#include virtual="/incs/fragReportContact.asp"-->

</table>
</form>


</body>
</html>
