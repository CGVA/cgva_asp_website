<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
	dim rsActiveEventData, rsActiveEventRows
	PAGE = Request("PAGE")
	EVENT_CD = Request("EVENT_CD")
	DIVISION_ID = Request("DIVISION_ID")

	If Request("CHOICE") <> "" Then
		CHOICE = Request("CHOICE")
		Session("CHOICE") = CHOICE
	Else
		CHOICE = Session("CHOICE")
	End If

	If Request("SCHEDULE_CHOICE") <> "" Then
		Session("SCHEDULE_CHOICE") = Request("SCHEDULE_CHOICE")
	Else
		Session("SCHEDULE_CHOICE")="REMAINING"
	End If

	''get all active events for the navigation bar
	SQL = 	"select EVENT_CD, EVENT_SHORT_DESC,LOCATION_DESC,EVENT_IMG from EVENT_TBL a " & _
			"LEFT JOIN LOCATION_TBL b ON a.LOCATION_CD = b.LOCATION_CD " & _
			"WHERE ACTIVE_EVENT_IND='Y' " & _
			"ORDER BY EVENT_CD DESC"
	set rs = cn.Execute(SQL)
	if not rs.EOF then
		rsActiveEventData = rs.GetRows
		rsActiveEventRows = UBound(rsActiveEventData,2)
	else
		rsActiveEventRows = -1
	end if


	If EVENT_CD <> "" Then
		SQL = 	"SELECT DIVISION_ID, DIVISION_CD,DIVISION_IMG FROM DIVISION_TBL " & _
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

	var cSchedule_0 = new Image();
	cSchedule_0.src = "images/schedule_0.gif";

	var cSchedule_1 = new Image();
	cSchedule_1.src = "images/schedule_1.gif";

	var cSchedule_2 = new Image();
	cSchedule_2.src = "images/schedule_2.gif";

	var cStandings_0 = new Image();
	cStandings_0.src = "images/standings_0.gif";

	var cStandings_1 = new Image();
	cStandings_1.src = "images/standings_1.gif";

	var cStandings_2 = new Image();
	cStandings_2.src = "images/standings_2.gif";

	var cTeams_0 = new Image();
	cTeams_0.src = "images/teams_0.gif";

	var cTeams_1 = new Image();
	cTeams_1.src = "images/teams_1.gif";

	var cTeams_2 = new Image();
	cTeams_2.src = "images/teams_2.gif";

	//tournament graphic
	var cTournament_0 = new Image();
	cTournament_0.src = "images/tournament_0.gif";

	var cTournament_1 = new Image();
	cTournament_1.src = "images/tournament_1.gif";

	var cTournament_2 = new Image();
	cTournament_2.src = "images/tournament_2.gif";

	/*Dynamic event graphics*/
	<%
	''event graphics
	For i = 0 to rsActiveEventRows

		rw("var v" & Replace(rsActiveEventData(0,i)," ","") & " = new Image();" & vbCrLf)
		rw("v" & Replace(rsActiveEventData(0,i)," ","") & ".src = ""images/" & rsActiveEventData(3,i) & "_0.gif"";" & vbCrLf & vbCrLf)

		rw("var v" & Replace(rsActiveEventData(0,i)," ","") & "_on = new Image();" & vbCrLf)
		rw("v" & Replace(rsActiveEventData(0,i)," ","") & "_on.src = ""images/" & rsActiveEventData(3,i) & "_1.gif"";" & vbCrLf & vbCrLf)

		rw("var v" & Replace(rsActiveEventData(0,i)," ","") & "_off = new Image();" & vbCrLf)
		rw("v" & Replace(rsActiveEventData(0,i)," ","") & "_off.src = ""images/" & rsActiveEventData(3,i) & "_2.gif"";" & vbCrLf & vbCrLf & vbCrLf)

	Next

	''division graphics (up over down)
	For i = 0 to rsDataRows

		''rw("var d" & rsData(1,i) & " = new Image();" & vbCrLf)
		''rw("d" & rsData(1,i) & ".src = ""images/" & rsData(2,i) & "_up.gif"";" & vbCrLf & vbCrLf)

		rw("var d" & i & "_up = new Image();" & vbCrLf)
		rw("d" & i & "_up.src = ""images/" & rsData(2,i) & "_up.gif"";" & vbCrLf & vbCrLf)

		rw("var d" & i & "_over = new Image();" & vbCrLf)
		rw("d" & i & "_over.src = ""images/" & rsData(2,i) & "_over.gif"";" & vbCrLf & vbCrLf & vbCrLf)

		rw("var d" & i & "_down = new Image();" & vbCrLf)
		rw("d" & i & "_down.src = ""images/" & rsData(2,i) & "_down.gif"";" & vbCrLf & vbCrLf & vbCrLf)
	Next
	%>
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

	//function to do unpressed division graphic
	function imgUp(imgName){
	document[imgName].src=eval(imgName + "_up.src");
	//alert(document[imgName].src);
	}

	//function to do depressed division graphic
	function imgDown(imgName){
	document[imgName].src=eval(imgName + "_down.src");
	//alert(document[imgName].src);
	}

	//function to do rollover division graphic
	function imgRollover(imgName){
	document[imgName].src=eval(imgName + "_over.src");
	//alert(document[imgName].src);
	}

	//////////////////////////////////////////////////////////////
	//function to do unpressed schedule/standings/team graphic
	function imgUp0(imgName){
	document[imgName].src=eval(imgName + "_0.src");
	//alert(document[imgName].src);
	}

	//function to do rollover schedule/standings/team graphic
	function imgRollover1(imgName){
	document[imgName].src=eval(imgName + "_1.src");
	//alert(document[imgName].src);
	}

	//function to do depressed schedule/standings/team graphic
	function imgDown2(imgName){
	document[imgName].src=eval(imgName + "_2.src");
	//alert(document[imgName].src);
	}

	function updateScheduleChoice()
	{
		choice = FORM.SCHEDULE_CHOICE.options[FORM.SCHEDULE_CHOICE.selectedIndex].value;
		alert(choice);
		redirectLocation = "index2.asp?EVENT_CD=<%=EVENT_CD%>&DIVISION_ID=<%=DIVISION_ID%>&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" + choice;
		alert(redirectLocation);
		window.location = redirectLocation;
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
		<table cellpadding='0' cellspacing='0'>
		<tr>
		<%
			If rsDataRows > -1 Then

				''division graphics go here
				For i = 0 to rsDataRows
					''UPDATE THIS WITH DYNAMIC DIVISION GRAPHICS PULL
					If CStr(DIVISION_ID) = CStr(rsData(0,i)) then
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & rsData(0,i) & "'><img src='../images/" & rsData(2,i) & "_down.gif' border='0' name='d" & i & "' id='d" & i & "' alt='' onmouseover=""imgRollover('d" & i & "')"" onmouseout=""imgDown('d" & i & "')"" /></a></td>")
					Else
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & rsData(0,i) & "'><img src='../images/" & rsData(2,i) & "_up.gif' border='0' name='d" & i & "' id='d" & i & "' alt='' onmouseover=""imgRollover('d" & i & "')"" onmouseout=""imgUp('d" & i & "')"" /></a></td>")
					End If

				Next

				If DIVISION_ID <> "" Then
					''UPDATE THIS WITH DYNAMIC choice GRAPHICS PULL

					Select Case Session("CHOICE")

					Case "SCHEDULE",""
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'><img src='../images/schedule_2.gif' border='0' name='cSchedule' id='cSchedule' alt='' onmouseover=""imgRollover1('cSchedule')"" onmouseout=""imgDown2('cSchedule')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'><img src='../images/standings_0.gif' border='0' name='cStandings' id='cStandings' alt='' onmouseover=""imgRollover1('cStandings')"" onmouseout=""imgUp0('cStandings')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'><img src='../images/teams_0.gif' border='0' name='cTeams' id='cTeams' alt='' onmouseover=""imgRollover1('cTeams')"" onmouseout=""imgUp0('cTeams')"" /></a></td>")
						rw("<td><a class='menuBlack' target='_blank' href='tournament/" & EVENT_CD & "_" & DIVISION_ID & ".asp'><img src='../images/tournament_0.gif' border='0' name='cTournament' id='cTournament' alt='' onmouseover=""imgRollover1('cTournament')"" onmouseout=""imgUp0('cTournament')"" /></a></td>")
					Case "STANDINGS"
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'><img src='../images/schedule_0.gif' border='0' name='cSchedule' id='cSchedule' alt='' onmouseover=""imgRollover1('cSchedule')"" onmouseout=""imgUp0('cSchedule')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'><img src='../images/standings_2.gif' border='0' name='cStandings' id='cStandings' alt='' onmouseover=""imgRollover1('cStandings')"" onmouseout=""imgDown2('cStandings')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'><img src='../images/teams_0.gif' border='0' name='cTeams' id='cTeams' alt='' onmouseover=""imgRollover1('cTeams')"" onmouseout=""imgUp0('cTeams')"" /></a></td>")
						rw("<td><a class='menuBlack' target='_blank' href='tournament/" & EVENT_CD & "_" & DIVISION_ID & ".asp'><img src='../images/tournament_0.gif' border='0' name='cTournament' id='cTournament' alt='' onmouseover=""imgRollover1('cTournament')"" onmouseout=""imgUp0('cTournament')"" /></a></td>")

					Case "TEAMS"
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'><img src='../images/schedule_0.gif' border='0' name='cSchedule' id='cSchedule' alt='' onmouseover=""imgRollover1('cSchedule')"" onmouseout=""imgUp0('cSchedule')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'><img src='../images/standings_0.gif' border='0' name='cStandings' id='cStandings' alt='' onmouseover=""imgRollover1('cStandings')"" onmouseout=""imgUp0('cStandings')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'><img src='../images/teams_2.gif' border='0' name='cTeams' id='cTeams' alt='' onmouseover=""imgRollover1('cTeams')"" onmouseout=""imgDown2('cTeams')"" /></a></td>")
						rw("<td><a class='menuBlack' target='_blank' href='tournament/" & EVENT_CD & "_" & DIVISION_ID & ".asp'><img src='../images/tournament_0.gif' border='0' name='cTournament' id='cTournament' alt='' onmouseover=""imgRollover1('cTournament')"" onmouseout=""imgUp0('cTournament')"" /></a></td>")
					Case "TOURNAMENT"
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'><img src='../images/schedule_0.gif' border='0' name='cSchedule' id='cSchedule' alt='' onmouseover=""imgRollover1('cSchedule')"" onmouseout=""imgUp0('cSchedule')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'><img src='../images/standings_0.gif' border='0' name='cStandings' id='cStandings' alt='' onmouseover=""imgRollover1('cStandings')"" onmouseout=""imgUp0('cStandings')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'><img src='../images/teams_0.gif' border='0' name='cTeams' id='cTeams' alt='' onmouseover=""imgRollover1('cTeams')"" onmouseout=""imgUp0('cTeams')"" /></a></td>")
						rw("<td><a class='menuBlack' target='_blank' href='tournament/" & EVENT_CD & "_" & DIVISION_ID & ".asp'><img src='../images/tournament_0.gif' border='0' name='cTournament' id='cTournament' alt='' onmouseover=""imgRollover1('cTournament')"" onmouseout=""imgUp0('cTournament')"" /></a></td>")

					Case Else
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=SCHEDULE&SCHEDULE_CHOICE=" & Session("SCHEDULE_CHOICE") & "&SORT=TIME'><img src='../images/schedule_0.gif' border='0' name='cSchedule' id='cSchedule' alt='' onmouseover=""imgRollover1('cSchedule')"" onmouseout=""imgUp0('cSchedule')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=STANDINGS'><img src='../images/standings_0.gif' border='0' name='cStandings' id='cStandings' alt='' onmouseover=""imgRollover1('cStandings')"" onmouseout=""imgUp0('cStandings')"" /></a></td>")
						rw("<td><a class='menuBlack' href='index2.asp?EVENT_CD=" & EVENT_CD & "&DIVISION_ID=" & DIVISION_ID & "&CHOICE=TEAMS'><img src='../images/teams_0.gif' border='0' name='cTeams' id='cTeams' alt='' onmouseover=""imgRollover1('cTeams')"" onmouseout=""imgUp0('cTeams')"" /></a></td>")
						rw("<td><a class='menuBlack' target='_blank' href='tournament/" & EVENT_CD & "_" & DIVISION_ID & ".asp'><img src='../images/tournament_0.gif' border='0' name='cTournament' id='cTournament' alt='' onmouseover=""imgRollover1('cTournament')"" onmouseout=""imgUp0('cTournament')"" /></a></td>")

					End Select

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
			'<tr bgcolor="#FF3300">
			'<td height="20" class="style8"><'%'=EVENT_DESC & " @ " & LOCATION_DESC '%'></td>
			'</tr>
			%>
			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragEventInformation.asp" --></td>
			</tr>
		<%ElseIf Session("CHOICE") = "" Then
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if
		%>
			<tr bgcolor="#FF3300">
			<td height="20"><font class='cfontWhite14'><%=EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD%></font></td>
			</tr>


			<tr bgcolor="#FFFFFF">
			<td><!-- #include virtual="/incs/fragSchedule.asp" --></td>
			</tr>
		<%Else

			''SCHEDULE WILL BE THE DEFAULT PAGE
			Select Case Session("CHOICE")

			Case "SCHEDULE",""
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				rw("<tr bgcolor=""#FF3300"">")
				rw("<td height=""14""><font class='cfontWhite12'>" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Schedule</td>")
				rw("<td valign='bottom' align='right' height=""14"">")
					rw("<form name='FORM'><select name='SCHEDULE_CHOICE' onChange='updateScheduleChoice();'>")
					rw("<option value='REMAINING'")
						if Session("SCHEDULE_CHOICE") = "" or Session("SCHEDULE_CHOICE")="REMAINING" then rw(" selected") end if
						rw(">Remaining Weeks</option>")
					rw("<option value='ALL'")
						if Session("SCHEDULE_CHOICE") = "ALL" then rw(" selected") end if
						rw(">All Weeks</option>")
					rw("</select></form>")
				rw("</td>")
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

				SQL = 	"SELECT IsNull(w.WEEK_NUM,0) FROM WEEK_TBL w " & _
					"LEFT JOIN EVENT_TBL e " & _
					"ON w.WEEK_ID = e.WEEK_NUM_DISPLAY_IND " & _
					"WHERE e.EVENT_CD = '" & EVENT_CD & "'"

				''rw(SQL)
				set rs = cn.Execute(SQL)
				if not rs.EOF then
					if CInt(rs(0)) = 0 then
						rw("<tr bgcolor=""#FF3300"">")
						rw("<td height=""14""><font class='cfontWhite14'>" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Standings</td>")
						rw("</tr>")

					else
						rw("<tr bgcolor=""#FF3300"">")
						rw("<td height=""14""><font class='cfontWhite14'>" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Standings as of Week " & rs(0) & "</td>")
						rw("</tr>")

					end if
				end if

				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragStandings.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				''rw("</tr>")
			Case "TEAMS"
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				rw("<tr bgcolor=""#FF3300"">")
				rw("<td height=""14""><font class='cfontWhite14'>" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Teams</td>")
				rw("</tr>")

				rw("<tr>")
				rw("<td>")
%>
				<!-- #include virtual="/incs/fragTeams.asp" -->
<%
				rw("</td>")
				rw("</tr>")
				''rw("</tr>")
			Case "TOURNAMENT"
				SQL = 	"SELECT DIVISION_CD FROM DIVISION_TBL " & _
				"WHERE DIVISION_ID = '" & DIVISION_ID & "'"

				set rs = cn.Execute(SQL)
				if not rs.EOF then
					DIVISION_CD = rs(0)
				else
					DIVISION_CD = ""
				end if

				rw("<tr bgcolor=""#FF3300"">")
				rw("<td height=""14""><font class='cfontWhite14'>" & EVENT_DESC & " @ " & LOCATION_DESC & " - " & DIVISION_CD & " - Tournament</td>")
				rw("</tr>")

				rw("<tr>")
				rw("<td>")

				rw("</td>")
				rw("</tr>")
				''rw("</tr>")
			Case Else
				rw("<tr bgcolor=""#FF3300"">")
				rw("<td height=""14""><font class='cfontWhite14'>" & EVENT_DESC & " @ " & LOCATION_DESC & "</td>")
				rw("</tr>")

				rw("<tr bgcolor=""#FFFFFF"">")
				rw("<td>DEFAULT.</td>")
				rw("</tr>")
			End Select
		%>

		<%End If%>


	<%ElseIf PAGE <> "" Then%>
		<tr bgcolor="#FFFFFF">
		<td>
			<%
			''Pass the name of the file to the function.
			''Function getFileContents(strIncludeFile)
''			  Dim objFSO
''			  Dim objText
''			  Dim strPage


			  ''Instantiate the FileSystemObject Object.
			  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


			  ''Open the file and pass it to a TextStream Object (objText). The
			  ''"MapPath" function of the Server Object is used to get the
			  ''physical path for the file.
			  strIncludeFile = "incs/frag" & PAGE & ".asp"
			  ''rw(Server.MapPath(strIncludeFile))
			  Set objText = objFSO.OpenTextFile(Server.MapPath(strIncludeFile))


			  ''Read and return the contents of the file as a string.
			  ''getFileContents = objText.ReadAll
			  rw(objText.ReadAll)

			  objText.Close
			  Set objText = Nothing
			  Set objFSO = Nothing
			''End Function
			%>
		</td>
		</tr>

	<%Else%>

		<tr bgcolor="#FFFFFF">
		<td>
			<!-- #include virtual="/incs/fragMainPage.asp"-->
		</td>
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
