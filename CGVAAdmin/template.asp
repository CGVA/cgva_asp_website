<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	dim rsActiveEventData, rsActiveEventRows
	EVENT_CD = Request("EVENT_CD")
	DIVISION_ID = Request("DIVISION_ID")
	CHOICE = Request("CHOICE")

	''get all active events for the navigation bar
	SQL = 	"select EVENT_CD, EVENT_SHORT_DESC,LOCATION_DESC,EVENT_IMG from EVENT_TBL a " & _
			"LEFT JOIN LOCATION_TBL b ON a.LOCATION_CD = b.LOCATION_CD " & _
			"WHERE ACTIVE_EVENT_IND='Y' " & _
			"ORDER BY EVENT_SHORT_DESC, LOCATION_DESC"
	set rs = cn.Execute(SQL)
	if not rs.EOF then
		rsActiveEventData = rs.GetRows
		rsActiveEventRows = UBound(rsActiveEventData,2)
	else
		rsActiveEventRows = -1
	end if


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

	//image preloads for navigation links
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
	league_rules.src = "images/league rules_0.gif";

	var league_rules_on = new Image();
	league_rules_on.src = "images/league rules_1.gif";

	var league_rules_off = new Image();
	league_rules_off.src = "images/league rules_2.gif";


	var officiating = new Image();
	officiating.src = "images/officiating_0.gif";

	var officiating_on = new Image();
	officiating_on.src = "images/officiating_1.gif";

	var officiating_off = new Image();
	officiating_off.src = "images/officiating_2.gif";

	var last_dig = new Image();
	last_dig.src = "images/last dig_0.gif";

	var last_dig_on = new Image();
	last_dig_on.src = "images/last dig_1.gif";

	var last_dig_off = new Image();
	last_dig_off.src = "images/last dig_2.gif";

	var volleypalooza = new Image();
	volleypalooza.src = "images/volleypalooza_0.gif";

	var volleypalooza_on = new Image();
	volleypalooza_on.src = "images/volleypalooza_1.gif";

	var volleypalooza_off = new Image();
	volleypalooza_off.src = "images/volleypalooza_2.gif";


	var about_us = new Image();
	about_us.src = "images/about us_0.gif";

	var about_us_on = new Image();
	about_us_on.src = "images/about us_1.gif";

	var about_us_off = new Image();
	about_us_off.src = "images/about us_2.gif";


	var photos = new Image();
	photos.src = "images/photos_0.gif";

	var photos_on = new Image();
	photos_on.src = "images/photos_1.gif";

	var photos_off = new Image();
	photos_off.src = "images/photos_2.gif";


	var links = new Image();
	links.src = "images/links_0.gif";

	var links_on = new Image();
	links_on.src = "images/links_1.gif";

	var links_off = new Image();
	links_off.src = "images/links_2.gif";


	var matchmate = new Image();
	matchmate.src = "images/match mate_0.gif";

	var matchmate_on = new Image();
	matchmate_on.src = "images/match mate_1.gif";

	var matchmate_off = new Image();
	matchmate_off.src = "images/match mate_2.gif";


	var contact_us = new Image();
	contact_us.src = "images/contact us_0.gif";

	var contact_us_on = new Image();
	contact_us_on.src = "images/contact us_1.gif";

	var contact_us_off = new Image();
	contact_us_off.src = "images/contact us_2.gif";

	/*Dynamic event graphics*/
	<%For i = 0 to rsActiveEventRows

		rw("var v" & rsActiveEventData(0,i) & " = new Image();" & vbCrLf)
		rw("v" & rsActiveEventData(0,i) & ".src = ""images/" & rsActiveEventData(3,i) & "_0.gif"";" & vbCrLf & vbCrLf)

		rw("var v" & rsActiveEventData(0,i) & "_on = new Image();" & vbCrLf)
		rw("v" & rsActiveEventData(0,i) & "_on.src = ""images/" & rsActiveEventData(3,i) & "_1.gif"";" & vbCrLf & vbCrLf)

		rw("var v" & rsActiveEventData(0,i) & "_off = new Image();" & vbCrLf)
		rw("v" & rsActiveEventData(0,i) & "_off.src = ""images/" & rsActiveEventData(3,i) & "_2.gif"";" & vbCrLf & vbCrLf & vbCrLf)

	Next%>
	//+++++++++++++++++   functions   ++++++++++++++++++++
/*
		var x = new Image();
		x.src = "images/x_0.gif";

		var x_on = new Image();
		x_on.src = "images/x_1.gif";

		var x_off = new Image();
		x_off.src = "images/x_2.gif";
*/


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

				''division graphics go here
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

		<%If DIVISION_ID = "" Then%>
			<%
			'<tr bgcolor="#FF9900">
			'<td height="20" class="style8"><'%'=EVENT_DESC & " @ " & LOCATION_DESC '%'></td>
			'</tr>
			%>
			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragEventInformation.asp" --></td>
			</tr>
		<%ElseIf CHOICE = "" Then
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if
		%>
			<tr bgcolor="#FF9900">
			<td height="20" class="style8"><%=EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD%></td>
			</tr>

			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragDivisionInformation.asp" --></td>
			</tr>
		<%Else

			Select Case CHOICE

			Case "SCHEDULE"
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				rw("<tr bgcolor=""#FF9900"">")
				rw("<td height=""20"" class=""style8"">" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Schedule</td>")
				rw("</tr>")

				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragSchedule.asp" -->
<%
				rw("</td>")
				rw("</tr>")

			Case "STANDINGS"
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				SQL = 	"SELECT IsNull(WEEK_NUM_DISPLAY_IND,0) FROM EVENT_TBL " & _
				"WHERE EVENT_CD = '" & EVENT_CD & "'"

				set rs = cn.Execute(SQL)
				if rs(0) = 0 then
					rw("<tr bgcolor=""#FF9900"">")
					rw("<td height=""20"" class=""style8"">" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Standings</td>")
					rw("</tr>")

				else
					rw("<tr bgcolor=""#FF9900"">")
					rw("<td height=""20"" class=""style8"">" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Standings as of Week " & rs(0) & "</td>")
					rw("</tr>")

				end if


				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragStandings.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				rw("</tr>")
			Case "TEAMS"
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				rw("<tr bgcolor=""#FF9900"">")
				rw("<td height=""20"" class=""style8"">" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Teams</td>")
				rw("</tr>")

				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragTeams.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				rw("</tr>")
			Case Else
				rw("<tr bgcolor=""#FF9900"">")
				rw("<td height=""20"" class=""style8"">" & EVENT_DESC & " @ " & LOCATION_DESC & "</td>")
				rw("</tr>")

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
