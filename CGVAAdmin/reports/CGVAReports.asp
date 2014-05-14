<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

''	If Session("USER_ID") = "" then
''		Session("Err") = "Your session has timed out. Please log in again."
''		Response.Redirect("/logout.asp")
''	End If

%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - CGVA Reports</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderReportGraphics.asp" -->

<tr>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - CGVA Reports</u></b>
	</font>
	</div>

	<br />

	<table align='center'>
	<tr>
	<td>
		<a href='TeamsReport.asp' class='menuGray'>Teams</a>
		<br />
		<a href='FinancialReport.asp' class='menuGray'>Financial</a>
		<br />
		<a href='ContactsReport.asp' class='menuGray'>Contact Information</a>
		<br />
		<a href='RatingsReport.asp' class='menuGray'>Player Ratings</a>
		<br />
	</td>
	</tr>
	</table>


</td>
</tr>

<tr height='20'><td></td></tr>
<!--#include virtual = "/incs/fragReportContact.asp"-->
</table>

</body>
</html>


