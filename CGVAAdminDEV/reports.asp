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
	ElseIf Not Instr(Session("ACCESS"),"BOD") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

''	If Request.ServerVariables("SERVER_NAME") <> "xxxxxxxxx" then
''		svr = "NonProd"
''	Else
''		svr = "Prod"
''	End If

	sql = 	"SELECT EVENT_CD, " & _
			"EVENT_SHORT_DESC, " & _
			"EVENT_LONG_DESC " & _
			"FROM EVENT_TBL " & _
			"ORDER BY EVENT_SHORT_DESC, EVENT_LONG_DESC"

	set rs = cn.Execute(sql)

	if rs.EOF then
		rw("Error: No events are available from which to select.")
		Response.End
	else
		rsEventData = rs.GetRows
		rsEventRows = UBound(rsEventData,2)
	end if

%>


<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Reports</title>
<script language="JavaScript">

function checkVariables()
{
    if (document.rptVariables.TYPEOFREPORT.value == "")
    {
		alert("Please select a report.");
		document.rptVariables.TYPEOFREPORT.focus();
		return false;
    }

	var dateStart = document.rptVariables.SDATE.value;
	var dateStop = document.rptVariables.EDATE.value;

	var sDateArray1 = dateStart.split("/");
	var eDateArray1 = dateStop.split("/");

	var sMonth = parseInt(sDateArray1[0],10);
	var sDay = parseInt(sDateArray1[1],10);

	var eMonth = parseInt(eDateArray1[0],10);
	var eDay = parseInt(eDateArray1[1],10);

	//change 2 char year to 4 char year
	if(parseInt(sDateArray1[2],10) < 2000)
	{
		var sYear = 2000 + parseInt(sDateArray1[2],10);
		document.rptVariables.SDATE.value = sMonth + "/" + sDay + "/" + sYear;
		dateStart = document.rptVariables.SDATE.value;
	}
	else
	{
		var sYear = parseInt(sDateArray1[2],10);
	}


	//change 2 char year to 4 char year
	if(parseInt(eDateArray1[2],10) < 2000)
	{
		var eYear = 2000 + parseInt(eDateArray1[2],10);
		document.rptVariables.EDATE.value = eMonth + "/" + eDay + "/" + eYear;
		dateStop = document.rptVariables.EDATE.value;
	}
	else
	{
		var eYear = parseInt(eDateArray1[2],10);
	}


	if(document.rptVariables.DATETYPE[0].checked)
	{

		if(checkIt(dateStart,"Start Date"))
		{

			if(checkIt(dateStop,"End Date"))
			{
				checkDates();
			}

			else
			{
				document.rptVariables.EDATE.focus();
				return false;
			}

		}
		else
		{
			document.rptVariables.SDATE.focus();
			return false;
		}

	}

	return true;
}

//*************************************************************

function checkDates(){

	var sDateString = document.rptVariables.SDATE.value;
	var eDateString = document.rptVariables.EDATE.value;

	var sDateArray = sDateString.split("/");
	var eDateArray = eDateString.split("/");

	var sMonth = parseInt(sDateArray[0],10);
	var sDay = parseInt(sDateArray[1],10);
	var sYear = 2000 + parseInt(sDateArray[2],10);

	var eMonth = parseInt(eDateArray[0],10);
	var eDay = parseInt(eDateArray[1],10);
	var eYear = 2000 + parseInt(eDateArray[2],10);

	var sDate = -1 *(Date.UTC(sYear,sMonth,sDay));
	var eDate = -1 *(Date.UTC(eYear,eMonth,eDay));

	//eDate = Math.abs(eDate);
	//sDate = Math.abs(sDate);

	//alert(sDateString);
	//alert(eDateString);
	if (Date.parse(sDateString) > Date.parse(eDateString)){

		alert("The Start Date entered must be less than the End Date");
		document.rptVariables.EDATE.focus();
		event.returnValue = false;
	}
	else
	{
		return true;
	}

}

//*************************************************************

function checkIt(dateStr,dateType) {

	var datePat = /^(\d{1,2})(\/|-)(\d{1,2})\2(\d{4})$/;
	var matchArray = dateStr.match(datePat); // is the format ok?

	if (matchArray == null) {
		alert("The " + dateType + " date " + dateStr + " is not a valid format.")
		return false;
	}

	month = matchArray[1]; // parse date into variables
	day = matchArray[3];
	year = matchArray[4];

	if (month < 1 || month > 12) { // check month range
		alert("The " + dateType + " Month must be between 1 and 12.");
		return false;
	}

	if (day < 1 || day > 31) {
		alert("The " + dateType + " Day must be between 1 and 31.");
		return false;
	}

	if ((month==4 || month==6 || month==9 || month==11) && day==31) {
		alert("The " + dateType + " Month " + month + " doesn't have 31 days.")
		return false;
	}

	if (month == 2) { // check for february 29th
		var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
		if (day>29 || (day==29 && !isleap)) {
			alert("February " + year + " doesn't have " + day + " days.");
			return false;
		}
	}

	return true;  // date is valid
}

</script>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->


<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont14'>
	<b><u>CGVA - Reports</u></b>
	</font>
	<br /><br />

	<form method="POST" action="reports_report.asp" id="rptVariables" name="rptVariables" onSubmit="return checkVariables();">

	<!-- BEGIN MAIN BODY (2 panes - left & right) -->
	<table align='center' border="0" cellpadding="0" cellspacing="0">

	<tr>
	<!-- BEGIN LEFT BODY CONTENT  -->
	<td align="left" valign="top">

		<!-- BEGIN reportVariables -->
		<table width="100%" border="0" cellpadding="0" cellspacing="0">

		<tr>
		<td height="8"><img src="../images/spacer.gif" width="1" height="8"></td>
		</tr>

		<tr>
		<td align="center" valign="middle"><font class='cfont10'><b>&nbsp;Select the report you wish to view, select valid filtering options, and click submit.</b></font></td>
		</tr>

		<tr>
		<td align="center" valign="middle">
			<font class='cfont10'><b>
			<ul>
			<li>To obtain a list of all current CGVA members, run the Contact Information Report with no Event selected.</li>
			</ul>
			</b></font>
		</td>
		</tr>

		</table>


		<table width="100%" cellpadding="0" cellspacing="0" border="0">

		<tr>
		<td colspan='3' align="left" valign="middle"><img src="../images/spacer.gif" WIDTH="30" HEIGHT="20"></td>
		</tr>

		<tr>
		<td align="left" valign="top">

			<!-- REPORT VARIABLE DISPLAY FOR THE OLS SYSTEM -->
			<table width="100%" cellpadding="3" cellspacing="0" border="0">

			<!-- REPORT VARIABLE - REQUIRED FOR ALL REPORTS -->
			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../images/rpt_ReportOptions.gif" width="235" height="15"></td>
			</tr>

			<tr>
			<td width='1%'><font class='cfont10'>Report:</font></td>
			<td align='left'>
				<select name="TYPEOFREPORT" size="1">
				<option value=""<% If Session("TYPEOFREPORT") = "" Then rw(" SELECTED") %>>-select-</option>
				<option value="REGISTRATION"<% If Session("TYPEOFREPORT") = "REGISTRATION" Then rw(" SELECTED") %>>Registered/Paid players per Event</option>
				<option value="CONTACTS"<% If Session("TYPEOFREPORT") = "CONTACTS" Then rw(" SELECTED") %>>Contact Information</option>
				<option value="RATINGS"<% If Session("TYPEOFREPORT") = "RATINGS" Then rw(" SELECTED") %>>Player Ratings</option>
				<option value="FASTLANE"<% If Session("TYPEOFREPORT") = "FASTLANE" Then rw(" SELECTED") %>>Fast Lane Personnel</option>
				<option value="TEAM"<% If Session("TYPEOFREPORT") = "TEAM" Then rw(" SELECTED") %>>Team Information</option>
				<option value="RESULTS"<% If Session("TYPEOFREPORT") = "RESULTS" Then rw(" SELECTED") %>>Event Results</option>
				<option value="SCHEDULE"<% If Session("TYPEOFREPORT") = "SCHEDULE" Then rw(" SELECTED") %>>Event Schedule</option>
				<option value="xxx"<% If Session("TYPEOFREPORT") = "xxx" Then rw(" SELECTED") %>></option>
				</select>
			</td>
			</tr>

			<tr>
			<td width='1%'><font class='cfont10'>Event:</font></td>
			<td align='left'>
				<select name="EVENT" size="1">
				<option value=""<% If Session("EVENT") = "" Then rw(" SELECTED") %>>-select-</option>
				<% For i = 0 to rsEventRows%>
					<option value="<%=rsEventData(0,i)%>"<%If Session("EVENT") = rsEventData(0,i) Then rw(" SELECTED") %>><%=rsEventData(1,i) & "(" & rsEventData(2,i) & ")"%></option>
				<%Next%>
				</select>
			</td>
			</tr>

			<!-- END REPORT VARIABLE -->

			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../images/spacer.gif" HEIGHT="20"></td>
			</tr>

			<!-- SCROLLING, PRINTER FRIENDLY OR EXCEL DISPLAY -->
			<tr>
			<td colspan='2' align="left" valign="middle"><img src="../images/rpt_DisplayOptions.gif" width="235" height="15"></td>
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
		<td width="5"><img src="../images/spacer.gif" width="5" height="1"></td>
		<!-- END SPACER -->

		<!-- SCROLLING, PRINTER FRIENDLY OR EXCEL DISPLAY -->
		<td align="left" valign="top">

			<table width="100%" cellpadding="3" cellspacing="0" border="0">

			<tr>
			<td colspan='2' height="20" align="left" valign="middle"><img src="../images/rpt_DateRange.gif" width="235" height="15"></td>
			</tr>

			<tr>
			<td><input type="radio" value="daily" name="DATETYPE" <%If Session("DATETYPE")="" or Session("DATETYPE")="daily" then%> checked<%End If%>><font class='cfont10'>Start Date</font>&nbsp;</td>
			<td>
				<%
					If Session("SDATE") = "" Then
						sdate=dateadd("d",-datepart("d",date())+1,date())
				%>
						<input type="text" name="SDATE" value="<%=sdate%>" size="10">

				<%Else%>
					<input type="text" name="SDATE" value="<%=Session("SDATE")%>" size="10">
				<%End If%>

			</td>
			</tr>

			<tr>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class='cfont10'>End Date</font>&nbsp;</td>
			<td>
				<%If Session("EDATE") = "" Then%>
					<input type="text" name="EDATE" value="<%=Month(dateadd("d",0,date())) & "/" & Day(dateadd("d",0,date())) & "/" & Year(dateadd("d",0,date()))%>" size="10">
				<%Else%>
					<input type="text" name="EDATE" value="<%=Session("EDATE")%>" size="10">
				<%End If%>
			</td>
			</tr>

			<tr>
			<td><input type="radio" value="monthly" name="DATETYPE" <%If Session("DATETYPE")="monthly" then%> checked<%End If%>><font class='cfont10'>Monthly</font>&nbsp;</td>
			<td>
				<select name="MDATE" size="1">
				<%	begDate = "9/01/2007"
					curDate = date()
					difDate = DateDiff("m",begDate,curDate)
					i = 0

					Do until i > difDate
						If Session("MDATE") = "" and (DatePart("m",curDate) = DatePart("m",begDate) and DatePart("yyyy",curDate) = DatePart("yyyy",begDate)) Then
							rw("<option value='" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "' SELECTED>" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "</option>" & vbcrlf)
						ElseIf Session("MDATE") <> "" and DatePart("m",Session("MDATE")) = DatePart("m",begDate) and DatePart("yyyy",Session("MDATE")) = DatePart("yyyy",begDate) Then
							rw("<option value='" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "' SELECTED>" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "</option>" & vbcrlf)
						Else
							rw("<option value='" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "'>" & MonthName(DatePart("m",begDate),true) & " " & Datepart("yyyy",begDate) & "</option>" & vbcrlf)
						End If

						begDate = DateAdd("m",1,begDate)
						i = i + 1
						Loop
				%>
				</select>
			</td>
			</tr>

			<tr>
			<td><input type="radio" value="alldates" name="DATETYPE" <%If Session("DATETYPE")="alldates" then%> checked<%End If%>><font class='cfont10'>All Dates</font>&nbsp;</td>
			<td>&nbsp;</td>
			</tr>

			<tr>
			<td colspan="2"><font class='cfont10'><font color='gray'>Click radio button for Daily, Monthly or All Dates</font></font></td>
			</tr>
			<!-- END DATE VARIABLE -->

			<!-- SPACER -->
			<tr>
			<td colspan="2"><img src="../images/spacer.gif" height="10"></td>
			</tr>
			<!-- END SPACER -->

			</table>

		</td>
		</tr>

		<!-- SPACER -->
		<tr>
		<td colspan="3"><img src="../images/spacer.gif" height="10"></td>
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
		<td colspan="3" height="10"><img src="../images/spacer.gif" width="1" height="10"></td>
		</tr>

		<tr>
		<td width="25" height="1"><img src="../images/spacer.gif" width="25" height="1"></td>
		<td width="450" height="1" bgcolor="black"><img src="../images/spacer.gif" width="1" height="1"></td>
		<td width="25" height="1"><img src="../images/spacer.gif" width="25" height="1"></td>
		</tr>

		<tr>
		<td colspan="3" height="10"><img src="../images/spacer.gif" width="1" height="10"></td>
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
<!--#include virtual="/incs/fragContact.asp"-->

</table>
</form>


</body>
</html>
