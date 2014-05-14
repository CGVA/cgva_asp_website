<!-- #include virtual="/incs/rw.asp" -->

<%

	dim fieldArray(10)
	fieldArray(0) = "MATCH_ID"
	fieldArray(1) = "WEEK_ID"
	fieldArray(2) = "TIME_ID"
	fieldArray(3) = "COURT_NUM"
	fieldArray(4) = "TEAM1_TEAM_ID"
	fieldArray(5) = "TEAM2_TEAM_ID"
	fieldArray(6) = "REF_TEAM_ID"
	fieldArray(7) = "TEAM1_MVP_ID"
	fieldArray(8) = "TEAM2_MVP_ID"
	fieldArray(9) = "TEAM1_INCLUDE_DIV_STATS_IND"
	fieldArray(10) = "TEAM2_INCLUDE_DIV_STATS_IND"

	rw("<html><body>" & vbCrLf)
	rw("function disableCheck(obj)" & vbCrLf)
	rw("{" & vbCrLf & vbCrLf)
	rw("rowValue = obj.value;" & vbCrLf)
	rw("var idValue = eval(MODIFY.ID + rowValue);" & vbCrLf)

	For i = 0 to UBound(fieldArray)
		rw("var " & fieldArray(i) & " = eval(MODIFY." & fieldArray(i) & " + rowValue);" & vbCrLf)
	Next

	rw("if(idValue.checked)" & vbCrLf)
	rw("{" & vbCrLf & vbCrLf)

	For i = 0 to UBound(fieldArray)
		rw(fieldArray(i) & ".disabled = false;" & vbCrLf)
	Next

	rw("}else{" & vbCrLf & vbCrLf)

	For i = 0 to UBound(fieldArray)
		rw(fieldArray(i) & ".disabled = true;" & vbCrLf)
	Next

	rw("}" & vbCrLf & vbCrLf)
	rw("}" & vbCrLf)


	For i = 0 to UBound(fieldArray)
		rw("<tr>" & vbCrLf)
		rw("<td align='right' valign='middle'><font size='2' face='arial'><b>" & fieldArray(i) & "</b></font></td>" & vbCrLf)
		rw("<td><input type=""text"" name='" & fieldArray(i) & "' size='' maxlength=''></td>" & vbCrLf)
		rw("</tr>" & vbCrLf & vbCrLf)

		rw("<tr>" & vbCrLf)
		rw("<td align='right' valign='middle'><font size='2' face='arial'><b>" & fieldArray(i) & "</b></font></td>" & vbCrLf)
		rw("<td>" & vbCrLf)
		rw("<select name='" & fieldArray(i) & "'>" & vbCrLf)
		rw("" & vbCrLf)
		rw("For i = 0 to " & xxx & vbCrLf)
		rw("rw(<option value=' & xxx(0,i) & '> & xxx(1,i) & </option>)" & vbCrLf)
		rw("Next" & vbCrLf)
		rw("" & vbCrLf)
		rw("</select>" & vbCrLf)
		rw("</td>" & vbCrLf)
		rw("</tr>" & vbCrLf & vbCrLf)
	Next

	rw("</body></html>" & vbCrLf)
%>

