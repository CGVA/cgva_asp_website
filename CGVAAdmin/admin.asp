<%
	''rw("ACCESS:" & Session("ACCESS"))
	''Response.End

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("index.asp")
''	Else
''		rw("ACCESS:" & Session("ACCESS"))
''		Response.End
	End If
%>
<!--#include virtual="/incs/fragHeader.asp"-->
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor="#FFFFFF">
<td>
	<form name='admin'>
	Select an area to update:
	<%If Instr(Session("ACCESS"),"ADMIN") > 0 Then%>
		<select name='choice' onChange="changePage();">
		<option value='admin.asp'>-select-</option>
		<option value='Division.asp'>Divisions</option>
		<option value='Event.asp'>Events</option>
		<option value='EventType.asp'>Event Types</option>
		<option value='FirstContact.asp'>First Contacts</option>
		<option value='Location.asp'>Locations</option>
		<option value='MatchGame.asp'>Match Games</option>
		<option value='MatchSchedule.asp'>Match Schedule</option>
		<option value='Person.asp'>Personnel</option>
		<option value='PlayerRatings.asp'>Player Ratings</option>
		<option value='Rater.asp'>Raters</option>
		<option value='PlayerRefCertification.asp'>Ref Certifications</option>
		<option value='Registration.asp'>Registration</option>
		<option value='Team.asp'>Teams</option>
		<option value='TeamMember.asp'>Team Personnel</option>
		<option value='Time.asp'>Time</option>
		<option value='UserLogin.asp'>User Login</option>
		<option value='Week.asp'>Week</option>
		<option value='ActivePeriod.asp'>Active Period (for cgva.org display)</option>
		<option value='admin.asp'></option>
		<option value='RoleType.asp'>Access Role Types</option>
		<option value='PersonAccess.asp'>Person Access</option>
		<option value='upload.asp'>CGVA Documents</option>
		<option value='reports.asp'>Reports</option>
		<option value='admin.asp'></option>
		<option value='LastDigRegistration.asp'>Last Dig Registration</option>
		</select>
	<%ElseIf Instr(Session("ACCESS"),"EDIT") > 0 Then%>
		<select name='choice' onChange="changePage();">
		<option value='admin.asp'>-select-</option>
		<option value='Division.asp'>Divisions</option>
		<option value='Team.asp'>Teams</option>
		<option value='TeamMember.asp'>Team Personnel</option>
		<option value='MatchSchedule.asp'>Match Schedule</option>
		<option value='MatchGame.asp'>Match Games</option>
		<option value='ActivePeriod.asp'>Active Period (for cgva.org display)</option>
		<option value='admin.asp'></option>
		<option value='upload.asp'>CGVA Documents</option>
		<option value='reports.asp'>Reports</option>
		<option value='admin.asp'></option>
		<option value='LastDigRegistration.asp'>Last Dig Registration</option>
		</select>
	<%ElseIf Instr(Session("ACCESS"),"BOD") > 0 Then%>
		<select name='choice' onChange="changePage();">
		<option value='admin.asp'>-select-</option>
		<option value='upload.asp'>CGVA Documents</option>
		<option value='reports.asp'>Reports</option>
		</select>
	<%Else%>
		Response.Redirect("index.asp")
	<%End If%>
	</form>
</td>
</tr>

<%If Instr(Session("ACCESS"),"ADMIN") > 0 Then%>
	<tr bgcolor="#FFFFFF">
	<td>
		<table align='center'>
		<tr>
		<td>
			<font class='cfont10'>
			Set up order:<br />
			First Contacts<br />
			Locations<br />
			Event Types<br />
			People<br />
			Raters<br />
			Player Ratings<br />
			Events<br />
			Registrations<br />
			Weeks<br />
			Times<br />
			Divisions<br />
			Teams<br />
			Team Personnel<br />
			Match Schedule<br />
			Match Games<br />
			</font>
		</td>
		</tr>
		</table>



	</td>
	</tr>
<%End If%>

<!--#include virtual="/incs/fragContact.asp"-->

</table>
<script language='Javascript'>
<!--
	admin.choice.focus();

	function changePage()
	{
		choice = admin.choice.options[admin.choice.selectedIndex].value;
		//alert(choice);
		window.location = choice;
	}
//-->
</script>
</body>
</html>
