<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	EVENT_CD = Request("EVENT_CD")
	DIVISION_ID = Request("DIVISION_ID")
	CHOICE = Request("CHOICE")

	If EVENT_CD <> "" Then
		SQL = 	"SELECT DIVISION_ID, DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE EVENT_CD = '" & EVENT_CD & "' " & _
				"ORDER BY DIVISION_CD"

		set rs = cn.Execute(SQL)

		if rs.EOF then
			Response.Redirect("http://www.cgva.org")
		else
			rsData = rs.GetRows
			rsDataRows = UBound(rsData,2)
		end if

		SQL = 	"select EVENT_SHORT_DESC,LOCATION_DESC from EVENT_TBL a " & _
				"LEFT JOIN LOCATION_TBL b ON a.LOCATION_CD = b.LOCATION_CD " & _
				"WHERE EVENT_CD = '" & EVENT_CD & "'"
		set rs = cn.Execute(SQL)
		if not rs.EOF then
			EVENT_DESC = rs(0)
			LOCATION_DESC = rs(1)
		end if

	Else
		rsDataRows = -1
		EVENT_DESC = ""
		LOCATION_DESC = ""
	End If
%>

<!-- #include virtual="/incs/rw.asp" -->
<!--#include virtual="/incs/fragHeader.asp"-->
<!-- #include virtual="/incs/fragStyle.asp" -->

<!-- scripts can go here -->
<script>
<!--
<!--

	//+++++++++++++++++   image preloads   ++++++++++++++++

	//image preloads for bottomlinks
	var home = new Image();
	home.src = "images/home_0.gif";

	var home_on = new Image();
	home_on.src = "images/home_1.gif";

	var home_off = new Image();
	home_off.src = "images/home_2.gif";

	var directions = new Image();
	directions.src = "images/directions_0.gif";

	var directions_on = new Image();
	directions_on.src = "images/directions_1.gif";

	var directions_off = new Image();
	directions_off.src = "images/directions_0.gif";


	var league_rules = new Image();
	league_rules.src = "images/league%20rules_0.gif";

	var league_rules_on = new Image();
	league_rules.src = "images/league%20rules_1.gif";

	var league_rules_off = new Image();
	league_rules_off.src = "images/league%20rules_2.gif";


	/*
	var league_rules_on = new Image();
	league_rules.src = "images/league rules_1.gif";

	var league_rules_off = new Image();
	league_rules_off.src = "images/league rules_2.gif";

	"images/officiating_1.gif"
	"images/officiating_2.gif"
	"images/last dig_1.gif"
	"images/last dig_2.gif"
	"images/volleypalooza_1.gif"
	"images/volleypalooza_2.gif"
	"images/about us_1.gif"
	"images/about us_2.gif"
	"images/photos_1.gif"
	"images/photos_2.gif"
	"images/links_1.gif"
	"images/links_2.gif"
	"images/match mate_1.gif"
	"images/match mate_2.gif"
	"images/contact us_1.gif"
	"images/contact us_2.gif"
	*/
	//+++++++++++++++++   functions   ++++++++++++++++++++


	//function to do mouseover
	function imgOver(imgName){
	document[imgName].src=eval(imgName + "_on" + ".src");
	//alert(document[imgName].src);
	}

	//function to do deactivate
	function imgOut(imgName){
	document[imgName].src=eval(imgName + ".src");
	//alert(document[imgName].src);
	}
-->
</script>

</head>

<!-- #include virtual="/incs/header.asp" -->

<table bgcolor='#000066' width='780' align='center' border='0' cellpadding='0' cellspacing='0'>
<tr>
<td colspan='3'><!-- #include virtual="/incs/fragSiteHeaderGraphics.asp" --></td>
</tr>

<tr>
<td valign='top'><!-- #include virtual="/incs/fragNavigation.asp" --></td>
<td>&nbsp;</td>
<td valign='top' width='100%'>

	<table width='100%' bgcolor="#000066" border="0" cellpadding='0' cellspacing='0'>

	<tr>
	<td>
		<table cellpadding='4' cellspacing='1'>
		<tr>
		<%
			If rsDataRows > -1 Then

				For i = 0 to rsDataRows
					rw("<td bgcolor='#FFFF33'><font class='cfont10'><b>&nbsp;&nbsp;<a class='menuBlack' href='template.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & rsData(0,i) & "'>" & rsData(1,i) & "</a>&nbsp;&nbsp;</b></font></td>")
				Next

				If DIVISION_ID <> "" Then
					rw("<td bgcolor='#00FF00'><font class='cfont10'><b>&nbsp;&nbsp;<a class='menuBlack' href='template.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE'>Schedule</a>&nbsp;&nbsp;</b></font></td>")
					rw("<td bgcolor='#00FF00'><font class='cfont10'><b>&nbsp;&nbsp;<a class='menuBlack' href='template.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'>Standings</a>&nbsp;&nbsp;</b></font></td>")
					rw("<td bgcolor='#00FF00'><font class='cfont10'><b>&nbsp;&nbsp;<a class='menuBlack' href='template.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'>Teams</a>&nbsp;&nbsp;</b></font></td>")
				End If

			End If
		%>
		</tr>
		</table>
	</td>
	</tr>

	<%If EVENT_DESC <> "" Then%>
		<tr bgcolor="#FF9900">
		<td height="20" class="style8"><%=EVENT_DESC & " @ " & LOCATION_DESC %></td>
		</tr>

		<%If DIVISION_ID = "" Then%>
			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragEventInformation.asp" --></td>
			</tr>
		<%ElseIf CHOICE = "" Then%>
			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragDivisionInformation.asp" --></td>
			</tr>
		<%Else

			Select Case CHOICE

			Case "SCHEDULE"
				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragSchedule.asp" -->
<%
				rw("</td>")
				rw("</tr>")

			Case "STANDINGS"
				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragStandings.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				rw("</tr>")
			Case "TEAMS"
				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragTeams.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				rw("</tr>")
			Case Else
				rw("<tr bgcolor=""#FFFFFF"">")
				rw("<td>DEFAULT.</td>")
				rw("</tr>")
			End Select
		%>

		<%End If%>


	<%Else%>
		<tr bgcolor="#FF9900">
		<td height="20" class="style8">Latest News Etc Will display...</td>
		</tr>

		<tr bgcolor="#FFFFFF">
		<td>This section would include the default home page information.</td>
		</tr>

	<%End If%>
	<tr>
	<td height='20'></td>
	</tr>

	</table>

</td>
</tr>

</table>

<!-- #include virtual="/incs/fragSiteContact.asp" -->

<br /><br />
</body>
</html>
