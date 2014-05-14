<%@Language=VBScript %>
<% Option Explicit %>
<%
'JVI 2004-04-16 as we will be calling this page from a client application and the user from client
'cannot create session variable from the client application, we need this code.
'if (Request("called") = "dedup") Then
'	Session.SessionID = now()
'end if

'SAS 2009-11-03 - eliminated table borders and set consistent cellpadding=2
'Also centered date column in Patient Notes

if (Request("called") <> "dedup") Then
	Session.Timeout=30
	IF NOT Session("loggedIn") = TRUE THEN
		Response.Redirect "loginTimedOut.asp"
	END IF
end if

'JVI 2004-04-16 New code to make this form to works with the client application
'dedup. It will create all the sessions variables.
if (Request("called") = "dedup") Then
	dim sSQL
	dim rs
	
	'first get the patient information to fill the session variables
	sSQL = "SELECT * FROM app_user WHERE login_id ='" & Request("user") & "' "
	set rs= server.CreateObject("ADODB.Recordset")
	
	rs.Open sSQL,objConn,adOpenForwardOnly,adLockReadOnly
	
	if rs.EOF Then 'should never happen but just in case
		Session("userID") = 9999
		Session("clinic_ID") = 9999 
		Session("level7")= TRUE
	else
		Session("userID") = rs("user_id")
		Session("clinic_ID") = Trim(rs("expansion_int_1"))
		IF rs("access_level_code") >6 THEN
			Session("level7")= TRUE
		END IF
	end if
	
	'create the session logged in
	Session("loggedIn") = TRUE
	rs.Close
	set rs=nothing
end if

'JPC 7/10/10
If IsNumeric(Request("ID")) and Request("ID") > 0 then
    Session("patientID") = Request("ID")
Else
    Response.Write("ERROR: There is an issue displaying the patient information.")
    Response.End
End If

	
'THIS HAS TO DO WITH WARNING ABOUT DUPLICATE VACCINES BEING ENTERED.
'SEE giveVaccine3.asp TO SEE THE PROCESS
'THIS JUST RESETS THE SESSION IF THE PERSON CANCELS OUT OF A FEW SCREENS
IF Session("needToWarnAgain")="No" THEN
	Session("needToWarnAgain")="Yes"
END IF
%>
<html>
<head>
<title>Public Health Information System</title>
<link rel="stylesheet" type="text/css" href="Style.css">
<style>
.lbfw { width=305px; }
</style>

<script language="JScript" runat="server">
    function SortVBArray(arrVBArray) {
        return arrVBArray.toArray().sort().join('\b');
    }
</script>

</head>
<body bgcolor="#FFFFFF">

<%

dim ageYear, ageMonth, monthsAlive, ageDays, sYear, sMonth, sDay
dim dayOfBirth, todaysDay

call Main()
Sub CalculateAge(birthDate,TodaysDate)
'=============================================================================
'JVI 2004-03-03: We need to show the age of the patient. 
'Here is how DateDiff function works. If we use "yyyy" it just look at the year
'field and return the difference. For example, Dec, 31, 1999 - Jan 1, 2000. It 
'returns 1 year. The same happens when compare months. For example, if we compare
'Feb 28 - March 1, dateDiff returns 1 month. So in order to fix that problem
'I will use month to get years and then compare the date (day) to fix for the 
'extra month that the DateDiff might return.
'=============================================================================

'Assume everybody is a new born
ageMonth =0
ageYear = 0
ageDays = 0
		
'Lets get the day of the months (birth date and Todays date)
dayOfBirth = Day(birthDate)
todaysDay = Day(TodaysDate)
		
'Calculate the Years
monthsAlive = DateDiff("m", birthDate, TodaysDate)

'If patient doesn't even is a month old then get out of here
if (monthsAlive = 0) Then
	ageDays=DateDiff("d", birthDate, TodaysDate)
	exit sub
end if

'Get number of years
if (monthsAlive >= 12) Then
	ageYear = int(monthsAlive / 12)
end if
		
'Calculate the number of months
ageMonth = (monthsAlive mod 12)
		
'Now we need to adjust the months and year value to avoid the rounding
if (ageMonth = 0) Then
	if (dayOfBirth > todaysDay) Then
		ageYear = ageYear - 1 'adjust year as we haven't complete a year yet even if tomorrow will be the year			
	end if
else		
	if (dayOfBirth > todaysDay) Then
		ageMonth = ageMonth - 1 'adjust month 			
	end if
end if

'Lets check if after adjustment we have 0 month and 0 year
if ((ageMonth = 0) and (ageYear = 0)) Then
	ageDays=DateDiff("d", birthDate, TodaysDate)
	exit sub
end if
		
'This is just for display - singular or plural
if (ageYear > 1) Then
	sYear = " years"
else
	sYear = " year"
end if
		
'This is just for display - singular or plural
if (ageMonth > 1) Then
	sMonth = " months) "
else 
	sMonth = " month) "
end if

End Sub

function getHealthPlan()
'Get the health plan name for display
'If the patient have one plan display the name of the plan, if he/she has more than one
'plan then display "Multiple".
dim sHealthName, rsHealth, sqlHealth

sqlHealth = "Select entry_desc from iz_multiple_entry where patient_id_system = '" & session("patientID") &"' " & _
	"and field_code = 3 "
set rsHealth = server.CreateObject("ADODB.Recordset")
rsHealth.CursorLocation = adUseClient
rsHealth.Open sqlhealth,objConn,adOpenStatic,adLockOptimistic

if rsHealth.EOF Then
	sHealthName = "&nbsp;"
else
	if (rsHealth.RecordCount = 1) Then
		sHealthName = rsHealth("entry_desc")
	else
		sHealthName = "Multiple"
	end if
end if

rsHealth.Close
set rsHealth=nothing

getHealthPlan = sHealthName

end function

Sub Main()
dim strSQL
strSQL = "SELECT * FROM patient WHERE patient_id_system ='" & Session("patientID") &"' "

dim objRS
set objRS=Server.CreateObject("Adodb.RecordSet")
objRS.Open strSQL, objConn

'SET SESSIONS VARIABLES FOR FIRST AND LAST NAME AND DOB SINCE THESE WILL BE DISPLAYED ON ALL PATIENT SCREENS
Session("firstName")= RTrim(objRS("first_name"))
Session("lastName")= RTrim(objRS("last_name"))
Session("dateOfBirth")= RTrim(objRS("date_of_birth"))

%>

<div align="center">
<table border="0" width="95%">
<tr><td colspan="2" align="center"><h3>PATIENT SUMMARY</h3></td></tr>
<tr><td colspan="2">

<%

dim strSQL7
strSQL7 = "SELECT * FROM app_user WHERE user_id ='" & Session("userID") &"' "

dim objRS7
set objRS7=Server.CreateObject("Adodb.RecordSet")
objRS7.Open strSQL7, objConn

dim strSQL8
strSQL8 = "SELECT * FROM clinic WHERE clinic_id ='" & objRS7("expansion_int_1") &"' "

dim objRS8
set objRS8=Server.CreateObject("Adodb.RecordSet")
objRS8.Open strSQL8, objConn
%>
	<div align="center">
	<p><b>USER NAME:&nbsp;</b><%= objRS7("user_first_name") %>&nbsp;<%= objRS7("user_last_name") %>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<b>CLINIC:&nbsp;</b><%= objRS8("clinic_name") %></p>
	<%
	objRS7.Close
	set objRS7=nothing
	objRS8.Close
	set objRS8=nothing
	%>
</tr>

<tr>
<td colspan="2">&nbsp;</td>
</tr>

<tr>
<td width="50%" valign="top">

	<div align="center">
	<table border="0" width="315" cellspacing="1" cellpadding="2">
	<tr>
	<td width="100%" align="center" colspan="2" background="images/formBackground.jpg">
	<p align="center"><b><font color="#FFFFFF">PATIENT INFORMATION</font></b></p>
	</td>
	</tr>
	
	<tr>
	<td align="right" valign="top" width="30%"><p><b>PATIENT ID:</b></p>
	</td>
	<td valign="top" width="70%"><p><%= RTrim(Left(objRS("patient_id_system"),3)) %>-<%= RTrim(Mid(objRS("patient_id_system"),4,5)) %>-<%= RTrim(Right(objRS("patient_id_system"),7)) %></p></td>
	</tr>
	
	<tr>
	<td align="right" valign="top"><p><b>NAME:</b></p>
	</td>
	<td valign="top" width="75%"><p><%= RTrim(objRS("first_name")) & "&nbsp;" & objRS("last_name") %></p></td>
	</tr>
		
	<tr>
	<td align="right" valign="top"><p><b>ALIAS:</b></p>
	</td>
	
	<td valign="top" width="75%"><p>
	<%
	'CHECK AND SEE IF PATIENT HAS AN ALIAS
	IF NOT RTrim(objRS("alias"))="" THEN
	Response.Write  RTrim(objRS("alias")) & "</p></td>"
	ELSE
	Response.Write "&nbsp;</p></td>"
	END IF
	%>
	</tr>
	
	<tr>
	<td align="right"><p><b>DOB:</b></p>
	</td>
	<td width="75%"><p>
	<%

	IF NOT RTrim(objRS("date_of_birth"))="" THEN
		call CalculateAge(Trim(objRS("date_of_birth")),Date())
		if ((ageDays = 0) and (ageMonth = 0) and (ageYear = 0)) Then
			Response.Write  RTrim(objRS("date_of_birth")) & " (0 days)</p></td>"
		else
			if (ageDays > 0) Then
				'This is just for display - singular or plural
				if (ageDays > 1) Then
					sDay = " days) "
				else
					sDay = " day) "
				end if
				Response.Write  RTrim(objRS("date_of_birth")) & " (" & ageDays & sDay & "</p></td>"
			else
				if ((ageYear = 0) and (ageMonth > 0)) Then
					Response.Write  RTrim(objRS("date_of_birth")) & " (" & ageMonth & sMonth & "</p></td>"
				else
					if ((ageYear > 0) and (ageMonth = 0)) Then
						Response.Write  RTrim(objRS("date_of_birth")) & " (" & ageYear & sYear & ")</p></td>"
					else
						Response.Write  RTrim(objRS("date_of_birth")) & " (" & ageYear & sYear & ", " & ageMonth & sMonth & "</p></td>"
					end if
				end if
			end if
		end if
	ELSE
		Response.Write "&nbsp;</p></td>"
	END IF
	%>
	</tr>
	
	<tr>
	<td align="right"><p><b>GENDER:</b></td>
	<td width="75%"><p>
	<% 
	IF objRS("gender_code") =1 THEN
	Response.Write "Female"
	ElseIf objRS("gender_code") =2 THEN
	Response.Write "Male"	
	Else
	Response.Write "&nbsp;"	
	END IF	
	%>
	</p>
	</td>
	</tr>

	<tr>
	<td align="right" nowrap><p><b>HEALTH PLAN:</b></td>
	<td width="75%"><p>
	<%
	Response.Write getHealthPlan() & "</p></td>"
	
	%>	
	</tr>
	
	<tr>
	<td align="right"><p><b>VFC:</b></td>
	<td width="75%"><p>
	<%
	IF NOT RTrim(objRS("vfc_eligibility_desc"))="" THEN
		Response.Write  RTrim(objRS("vfc_eligibility_desc")) & "</p></td>"
	ELSE
		Response.Write "&nbsp;</p></td>"
	END IF
	%>	
	</tr>
	
	<%
	IF NOT trim(objRS("patient_status"))="" THEN
		dim strSQL4
		strSQL4 = "SELECT code_desc FROM code WHERE code_group=14 and code=" & Trim(objRS("patient_status")) &" "

		dim objRS4
		set objRS4=Server.CreateObject("Adodb.RecordSet")
		objRS4.Open strSQL4, objConn
	END IF	
	%>
	
	
	<tr>
	<td align="right"><p><b>STATUS:</b></td>
	<td width="75%"><p>
	<%
	IF NOT trim(objRS("patient_status"))="" THEN
		IF NOT RTrim(objRS4("code_desc"))="" THEN
			Response.Write  RTrim(objRS4("code_desc")) & "</p></td>"
		ELSE
			Response.Write "&nbsp;</p></td>"
		END IF
		objRS4.Close
		set objRS4=nothing
	END IF
	%>		
	</tr>

	<tr>
	<td align="right"><p><b>CHART#:</b></p>
	</td>
	<td width="75%"><p>
	<%
	IF NOT RTrim(objRS("patient_id_local"))="" THEN
		Response.Write  RTrim(objRS("patient_id_local")) & "</p></td>"
	ELSE
		Response.Write "&nbsp;</p></td>"
	END IF
	%>			
	</tr>

	<tr>
	<%
	'JPC 7/10/10
	Response.Write "<td colspan='2' align='center'><p align='center'><a href='editPatientInfo.asp?ID=" & Server.HTMLEncode(Session("patientID")) &"'><img border='0' src='images/editButton.jpg' alt='editButton.jpg'></a></p></td>"
	%>
	</tr>
	</table>
	</div>

<br>

	<div align="center">
	<table border="0" width="315" cellspacing="1" cellpadding="2">
	<tr>
	<td width="100%" align="right" colspan="5" background="images/formBackground.jpg">
	<p align="center"><b><font color="#FFFFFF">VACCINE HISTORY</font></b></p>
	</td>
	</tr>
	<%
	dim strSQL3
	'strSQL3 = "SELECT * FROM vaccine_history WHERE patient_id_system ='" & Session("patientID") &"' ORDER BY vaccine_abbr, vaccine_id, dose_number"

	strSQL3 = strSQL3 & "SELECT a.adverse_reaction_desc, a.vaccine_id, a.vaccine_abbr, a.dose_number, convert(varchar(10), a.vaccination_date, 101) as [vaccination_date_formatted], (case when adverse_reaction_code = 1 then 'N' when adverse_reaction_code is null then 'N' when adverse_reaction_code = '' then 'N' else 'Y' end), left(b.code_desc,1) AS codeDesc "
	strSQL3 = strSQL3 & "FROM vaccine_history a, code b "
	strSQL3 = strSQL3 & "WHERE a.patient_id_system = '" & Session("patientID") &"' "
	strSQL3 = strSQL3 & "and b.code_group = 65 "
	strSQL3 = strSQL3 & "and b.code = isnull((select c.entry_code from iz_vaccine_history c "
	strSQL3 = strSQL3 & "where c.vaccination_id = a.expansion_int_2),9) "
	strSQL3 = strSQL3 & "ORDER BY vaccine_abbr, vaccine_id, dose_number "
	
	dim objRS3
	set objRS3=Server.CreateObject("Adodb.RecordSet")
	objRS3.Open strSQL3, objConn

	Response.Write "<tr><td align='left' valign='top' background='images/formBackground2.jpg'><p><b><font color='#FFFFFF'>VACCINE</font></b></p></td>"
	Response.Write "<td align='center' valign='top' background='images/formBackground2.jpg'><p><b><font color='#FFFFFF'>#</font></b></p></td>"
	Response.Write "<td align='center' valign='top' background='images/formBackground2.jpg'><p><b><font color='#FFFFFF'>DATE</font></b></p></td>"
	Response.Write "<td align='center' valign='top' background='images/formBackground2.jpg'><p><b><font color='#FFFFFF'>ADV</font></b></p></td>"	
	Response.Write "<td align='center' valign='top' background='images/formBackground2.jpg'><p><b><font color='#FFFFFF'>ENT</font></b></p></td>"		
	Response.Write "</tr>"
	
	IF objRS3.EOF THEN
		Response.Write "<tr><td align='center' valign='top' colspan='5'><p><i>NO VACCINES HAVE BEEN ADDED.</i></p></td></tr>"
	ELSE
		DO WHILE NOT objRS3.EOF
		Response.Write "<tr><td align='left' valign='top'><p><a href='editVaccination.asp?ID=" & request("ID") & "&vaccineID="& objRS3("vaccine_id") &"&dose="& objRS3("dose_number") &"'>" & objRS3("vaccine_abbr") & "</a></p></td>"
		Response.Write "<td align='center' valign='top'><p>" & objRS3("dose_number") & "</p></td>"
		Response.Write "<td align='center' valign='top'><p>" & objRS3("vaccination_date_formatted") & "</p></td>"
		IF ISNULL(Trim(objRS3("adverse_reaction_desc"))) OR Trim(objRS3("adverse_reaction_desc"))="Unknown" OR Trim(objRS3("adverse_reaction_desc"))="None" OR Trim(objRS3("adverse_reaction_desc"))="" THEN
			Response.Write "<td align='center' valign='top'><p>N</p></td>"
		ELSE
			Response.Write "<td align='center' valign='top'><p><font color='#FF0000'>Y</font></p></td>"
		END IF
		Response.Write "<td align='center' valign='top'><p>" & objRS3("codeDesc") & "</p></td>"
		Response.Write "</tr>"
		objRS3.MoveNext
		Loop
	END IF
	
	objRS3.Close
	set objRS3=nothing

	Response.Write "<tr>"
	'JPC 7/10/10
	Response.Write "<td align='center' valign='top' colspan='5'><p><a href='add_history.asp'><img src='images/addHistoryBtn2.jpg' border='0' alt='addHistoryBtn2.jpg'></a>&nbsp;&nbsp; <a href='giveVaccine.asp?ID=" &  Server.HTMLEncode(Session("patientID")) & "'><img src='images/giveVaccineBtn2.jpg' border='0' alt='giveVaccineBtn2.jpg'></a>&nbsp;&nbsp; "
	'JPC 7/10/10
	Response.Write "<a href='editVaccineHistory.asp?ID=" &  Server.HTMLEncode(Session("patientID")) &"'><img src='images/editHistoryBtn2.jpg' border='0' alt='editHistoryBtn2.jpg'></a></p></td>"
	%>
	</tr>
	
	<!--START RECOMMEND CODE-->	
	<form name="Vaccines" method="POST" action="recommendVaccine.asp">
	<tr>
	<td align="center" colspan="5"><select name="currentSerial" class="lbfw">
	<%
	dim sSQL01, rsGroup
	
	sSQL01 = "Select a.code_desc, b.vaccine_group_id from code a, (Select distinct vaccine_group_id from iz_vaccine_group) b  where b.vaccine_group_id = a.code and " & _
		"a.code_group = 47 Order by a.code_desc "
		
	set rsGroup = server.CreateObject("ADODB.Recordset")
	rsGroup.Open sSQL01,objConn,adOpenForwardOnly,adLockReadOnly
	
	'lets add the option ALL to the list box it will have a value of -1
	Response.Write "<option value=-1>All Series</option>" 
	
	do while not rsGroup.EOF
		Response.Write "<option value=" & rsGroup("vaccine_group_id")  & ">" & Trim(rsGroup("code_desc")) & "</option>" 
		rsGroup.MoveNext
	loop
	
	
	rsGroup.Close
	set rsGroup=nothing

	%>
	</select><br>

	<input src="images/recommendButton.jpg" type="Image" Alt="Recommend Vaccines" border="0" name="submit" WIDTH="124" HEIGHT="28">
	</td>
	</tr>
	</form>
	<!--END RECOMMEND CODE-->
	
	</table>
	</div>
</td>

<td width="50%" valign="top">
	<div align="center">
	<center>
	<table border="0" width="306" cellspacing="1" cellpadding="2">
		<tr>
		<td width="100%" colspan="2" align="center" background="images/formBackground.jpg">
		<p align="center"><b><font color="#FFFFFF">CONTACT INFORMATION</font></b></p>
		</td>
		</tr>
		
		<tr>
		<td align="right" valign="top"><p><b>PARENT:</b></p>
		</td>
		<%
		IF NOT Trim(objRS("mother_last_name"))="" THEN
		'THERE IS A NAME FOR MOTHER SO THIS IS WHO WE WILL SHOW IN THIS FIELD
			IF Rtrim(objRS("mother_middle_name"))="" THEN
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("mother_first_name")) &"&nbsp;"& RTrim(objRS("mother_last_name")) & "</p></td>"				
			ELSE
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("mother_first_name")) &"&nbsp;"& RTrim(objRS("mother_middle_name")) & "&nbsp;"& RTrim(objRS("mother_last_name")) & "</p></td>"
			END IF	
		END IF
		
		IF Trim(objRS("mother_last_name"))="" AND NOT Trim(objRS("father_last_name"))="" THEN
		'THERE IS NO VALUE FOR MOTHER SO WE WILL SHOW FATHER IN THIS FIELD
			IF Rtrim(objRS("father_middle_name"))="" THEN
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("father_first_name")) &"&nbsp;"& RTrim(objRS("father_last_name")) & "</p></td>"				
			ELSE
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("father_first_name")) &"&nbsp;"& RTrim(objRS("father_middle_name")) & "&nbsp;"& RTrim(objRS("father_last_name")) & "</p></td>"
			END IF	
		END IF
		
		IF NOT Trim(objRS("guardian_last_name"))="" AND (Trim(objRS("mother_last_name"))="" AND Trim(objRS("father_last_name"))="") THEN
		'THERE IS NO VALUE FOR MOTHER OR FATHER SO WE WILL SHOW GUARDIAN IN THIS FIELD
			IF Rtrim(objRS("guardian_middle_name"))="" THEN
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("guardian_first_name")) &"&nbsp;"& RTrim(objRS("guardian_last_name")) & "</p></td>"				
			ELSE
				Response.Write "<td width='75%'><p>" & Rtrim(objRS("guardian_first_name")) &"&nbsp;"& RTrim(objRS("guardian_middle_name")) & "&nbsp;"& RTrim(objRS("guardian_last_name")) & "</p></td>"
			END IF	
		END IF	
		
		IF Trim(objRS("mother_last_name"))="" AND Trim(objRS("father_last_name"))="" AND Trim(objRS("guardian_last_name"))=""THEN
		'THERE IS NO VALUE FOR MOTHER, FATHER OR GUARDIAN BUT WE STILL NEED TO DRAW THE BORDER FOR THE BOX
			Response.Write "<td width='75%'><p>&nbsp;</p></td>"
		END IF		
		%>
		</tr>
		
		<tr>
		<td align="right" valign="top"><p><b> ADDRESS:</b></p></td>
		<td width="75%"><p>
		<%
		Response.Write Trim(objRS("addr_line1"))
		IF NOT Trim(objRS("suite"))="" THEN
			Response.Write ", " & Trim(objRS("suite")) & "<br>"
		ELSE
			Response.Write "<br>"
		END IF
		%>
				
		<%
		IF NOT IsNull(objRS("addr_line2")) THEN
				IF NOT TRIM(objRS("addr_line2")) ="" THEN
					Response.Write Trim(objRS("addr_line2")) & "<br>"
				END IF	
		END IF	
		%>
		
		<%
		'LOOK AND FEEL STUFF
		IF NOT RTrim(objRS("city_desc"))= "Unknown" THEN
			Response.Write RTrim(objRS("city_desc"))  &", " 
		END IF
	
		IF NOT RTrim(objRS("state_desc"))= "Unknown" THEN
			Response.Write objRS("state_desc") & "&nbsp;"
		END IF
	
		IF NOT RTrim(objRS("zip")) ="" THEN
			Response.Write objRS("zip") 
		END IF
		%>
		</p>
		</td>
		</tr>
		
		<tr>
		<td align="right"><p><b>PRIMARY#:</b></p></td>
		<td width="75%"><p>
		<%
		'format the phone # to display in (XXX) XXX-XXXX format
		IF NOT TRIM(objRS("HOME_PHONE"))="" THEN
			Dim re
			Set re = New RegExp
  
			re.Pattern = "(\d{3})(\d{3})(\d{4})"
  
			Response.Write re.Replace(objRS("HOME_PHONE"), "($1) $2-$3")
  
			Set re = Nothing
		Else
			Response.Write "&nbsp;"
		END IF
		%>
		</p>
		</td>
		</tr>
									
		<!--		'REMOVED SECONDARY# AND REPLACED IT WITH PC (SEE BELOW)		<tr>		<td align="right"><p><b>SECONDARY#:</b></p>		</td>		<%			'Response.Write "<td width='75%'><p>"			'IF NOT TRIM(objRS("mother_work_phone"))="" THEN				'Dim re2				'Set re2 = New RegExp  				're2.Pattern = "(\d{3})(\d{3})(\d{4})"  				'Response.Write re2.Replace(objRS("mother_work_phone"), "($1) $2-$3")  				'Set re2 = Nothing			'Else				'Response.Write "&nbsp;"			'END IF								'Response.Write"</p></td>"						%>		</tr>		-->

		<tr>
		<td align="right"><p><b>PCP:</b></p>
		</td>
		<%
			IF (TRIM(objRS("provider_first_name"))="" OR ISNULL(objRS("provider_first_name"))) THEN
				Response.Write "<td width='75%'><p>&nbsp;"
			ELSE	
				Response.Write "<td width='75%'><p>" & TRIM(objRS("provider_first_name")) & "&nbsp;" & TRIM(objRS("provider_last_name")) & ",&nbsp;" & TRIM(objRS("provider_title_desc")) & " "
			END IF
			Response.Write"</p></td>"				
		%>
		</tr>





		
		<tr>
		<td align="right" valign="top"><p><b>REMINDER:</b></p></td>
		<td valign="top" width="75%"><p>		
		<%
		IF NOT RTrim(objRS("rem_letter_returned"))="" THEN
			'Response.Write  RTrim(objRS("rem_letter_returned")) & "</p></td>"
			Select Case RTrim(objRS("rem_letter_returned"))
			Case "0"
			Response.Write "None Sent</p></td>"
			Case "1"
			Response.Write "Pending</p></td>"
			Case "2"
			Response.Write "Returned</p></td>"
			Case "99"
			Response.Write "<p>Unknown</p></td>"
			Case Else
			Response.Write "&nbsp;</p></td>"
			End Select
		ELSE
			Response.Write "&nbsp;</p></td>"
		END IF
		%>	
		
		</tr>
		
		<tr>
		<%
		'JPC 7/10/10
		Response.Write "<td colspan='2' align='center'><p align='center'><a href='editGeneralInfo.asp?ID=" &  Server.HTMLEncode(Session("patientID")) &"'><img border='0' src='images/editButton.jpg' alt='editButton.jpg'></a></p></td>"
		%>
		</tr>
		</table>
		</center>
		</div>
<br>
	<div align="center">
	<table border="0" width="306" cellspacing="1" cellpadding="2">
		<tr>
		<td width="100%" align="right" colspan="3" background="images/formBackground.jpg">
		<p align="center"><b><font color="#FFFFFF">CONTRAINDICATIONS</font></b></p></td>
		</tr>
	
		<tr>
		<td align="center" colspan="3"><p>
		<%
		dim strSQL2
		strSQL2 = "SELECT * FROM code, patient_allergy WHERE patient_id_system ='" & Session("patientID") &"' AND code_group=19 AND allergy_code = code ORDER BY code_desc"
	
		dim objRS2
		set objRS2=Server.CreateObject("Adodb.RecordSet")
		objRS2.Open strSQL2, objConn
	
		Response.Write "<table align='left' width='100%'>"
		IF objRS2.EOF THEN
			Response.Write "<tr><td align='center'><p><i>NO CONTRAINDICATIONS ADDED.</i><br>"
		ELSE
			'Response.Write "<tr><td colspan='3' align='left'><p>"
			DO WHILE NOT objRS2.EOF
			Response.Write "<tr><td colspan='3' align='left'><p>"
			Response.Write  Trim(objRS2("code_desc")) & "</td></tr>"
			objRS2.MoveNext
			Loop
		END IF

		'JPC 7/10/10
		Response.Write "<tr><td colspan='3' align='center'><p><a href='editContraindications.asp?ID=" &  Server.HTMLEncode(Session("patientID")) &"'><img border='0' src='images/editButton.jpg' alt='editButton.jpg'></a></p></td></tr>"
		Response.Write "</table>"
	
		objRS2.close
		set objRS2=nothing
		%>
	</table>
	</div>
	
	<br>
	
	<div align="center">
	<table border="0" width="306" cellspacing="1" cellpadding="2">
		<tr>
		<td width="100%" align="right" colspan="3" background="images/formBackground.jpg">
		<p align="center"><b><font color="#FFFFFF">PATIENT NOTES</b></p></font></td>
		</tr>
	
		<tr>
			<td background="images/formBackground2.jpg"><p><b><font color="#FFFFFF">PRIORITY</font></b></p></td>
			<td background="images/formBackground2.jpg" align="center"><p><b><font color="#FFFFFF">DATE</font></b></p></td>
			<td background="images/formBackground2.jpg"><p><b><font color="#FFFFFF">NOTES</font></b></p></td>
		</tr>

			<%
			dim strSQL5
			strSQL5 = strSQL5 & "select Top 5 rtrim(b.code_desc) as codeDesc, convert(varchar(10), a.entry_date, 101) as [entry_date_formatted], substring(a.notes,1,60) as notes, patient_notes_id "
	  		strSQL5 = strSQL5 & "from iz_patient_notes a, code b "
	  		strSQL5 = strSQL5 & "where a.patient_id_system = '" & Session("patientID") &"' "
	  		strSQL5 = strSQL5 & "and b.code_group = 60 "
	   		strSQL5 = strSQL5 & "and b.code = a.priority_code "
			strSQL5 = strSQL5 & "order by a.priority_code, entry_date desc "

			dim objRS5
			set objRS5=Server.CreateObject("Adodb.RecordSet")
			'objRS5.Open strSQL5, objConn
			objRS5.CursorLocation = adUseClient
			objRS5.Open strSQL5, objConn, adOpenStatic, adLockOptimistic

			IF objRS5.EOF THEN
				Response.Write "<tr><td colspan='3'><p><i>There are no notes.</i></p></td></tr>"
			ELSE
				Do While Not objRS5.EOF
				Response.Write "<tr>"
				Response.Write "<td valign='top'><p>" & objRS5("codeDesc") & "</p></td>"
				Response.Write "<td valign='top' align='center'><p>" & objRS5("entry_date_formatted") & "</p></td>"
				Response.Write "<td valign='top'><p><a href='editNotes.asp?notesID=" & objRS5("patient_notes_id") & "'>" & objRS5("notes") & " "
				IF Len(objRS5("notes")) > 59 THEN
					Response.Write "...</a></p></td>"
				ELSE
					Response.Write "</a></p></td>"
				END IF
				Response.Write "</tr>"
				objRS5.MoveNext
				Loop
			END IF	

			IF objRS5.RecordCount >4 THEN
				Response.Write "<tr><td colspan='3' align='center'><p><i><a href='allNotes.asp'>See All Notes</a></i></p></td></tr>"
			END IF
			
			objRS5.close
			set objRS5=nothing
		%>
			</tr>		
			<tr>
			<td colspan="3" align="center">
			<%
			Response.Write "<a href='addPatientNotes.asp'><img src='images/addNewNote.jpg' alt='addNewNote.jpg' border='0'></a>"
			%>
			</td>
		</tr>
	</table>
	</div>

	<br>

	<div align="center">
	<table border="0" width="306" cellspacing="1" cellpadding="2">
		<tr>
		<td width="100%" align="right" colspan="3" background="images/formBackground.jpg">
		<p align="center"><b><font color="#FFFFFF">QUALIFYING INTERVIEW</b></p></font></td>
		</tr>
	
		<tr>
		<td colspan="3">
		
			<%
			dim strSQL6
			strSQL6 = strSQL6 & "select encounter_date, next_imm_date "
			strSQL6 = strSQL6 & "from iz_interview "
			strSQL6 = strSQL6 & "where patient_id_system = '" & Session("patientID") &"' "
			strSQL6 = strSQL6 & "order by encounter_date desc "
			
			dim objRS6
			set objRS6=Server.CreateObject("Adodb.RecordSet")
			objRS6.Open strSQL6, objConn

			IF objRS6.EOF THEN
				Response.Write "<tr><td colspan='3'><p><i>There are no qualifying interviews.</i></p></td></tr>"
			ELSE
				Response.Write "<p>"
				Do While Not objRS6.EOF
				Response.Write "<a href='editQualifyingInterview.asp?date=" & objRS6("encounter_date") &"&next_imm_date=" & objRS6("next_imm_date") &"'>" & objRS6("encounter_date") &"</a>, "
				objRS6.MoveNext
				Loop
				Response.Write "</p>"				
			END IF	
			
			objRS6.close
			set objRS6=nothing

		objRS.close
		set objRS=nothing
	end sub
		%>
			</td>
			</tr>
		
		</td>
		</tr>
	</table>
	</div>	

</td>
</tr>
</table>
</div>

<p align="center">
<%'JPC 7/10/10 %>
<a href="deletePatient.asp?ID=<%=  Server.HTMLEncode(Session("patientID")) %>"><img src="images/deletePatientBtn.jpg" border="0" alt="deletePatientBtn.jpg" WIDTH="124" HEIGHT="28"></a> &nbsp;&nbsp;
<a href="searchPatients.asp"><img src="images/doneButton.jpg" border="0" alt="doneButton.jpg" WIDTH="124" HEIGHT="28"></a></p>

</body>
</html>

















