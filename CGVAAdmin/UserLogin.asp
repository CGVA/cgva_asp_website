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
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	Dim rsPersonData,rsPersonRows, rsAvailablePersonData, rsAvailablePersonRows
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - User Login Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - User Login Administration</u></b>
	</font>
	</div>

	<br />

<%

class userlogin
''private segment, org


sub class_initialize()
	if request.totalbytes=0 then

	else

		if request("btnsave") <> "" then

			dim i

			''deal with deletes
			for each i in request.form

				if i = "REMOVE" then
					arrUserID = split(request("REMOVE"), ", ")

					for j = 0 to UBound(arrUserID)

						SQL = 	"DELETE FROM USER_LOGIN_TBL " & _
								"WHERE PERSON_ID='" & arrUserID(j) & "'"
						''rw(request("REMOVE"))
						''rw(SQL)
						cn.Execute(SQL)

					next

				end if

			next

			for each i in request.form
				''rw(i & "<br />")
				arrid = split(i, "_")
				''rw(arrid(0) & "<br />")
				''rw(arrid(1) & "<br />")
				''rw(arrid(2) & "<br />")

				if arrid(0) = "USERNAME" then

					sql = 	"update USER_LOGIN_TBL " & _
							"set [USERNAME] = '" & replace(request("USERNAME_" & arrid(1)),"'","''") & "', " & _
							"[PASSWORD] = '" & replace(request("PASSWORD_" & arrid(1)),"'","''") & "', " & _
							"[SECURITY_QUESTION] = '" & replace(request("SECURITYQUESTION_" & arrid(1)),"'","''") & "', " & _
							"REMINDER = '" & replace(request("REMINDER_" & arrid(1)),"'","''") & "', " & _
							"ANSWER = '" & replace(request("ANSWER_" & arrid(1)),"'","''") & "' " & _
							"where PERSON_ID = '" & arrid(1) & "'"
					''response.write sql & "<br />"
					cn.Execute(sql)

				end if

			next 'i

			session("message") = session("message") & "User Login information has been updated successfully.<br />"

			if request("PERSONIDnew") <> "" then

				sql = 	"insert into USER_LOGIN_TBL" & _
						"([PERSON_ID], " & _
						"[USERNAME], " & _
						"[PASSWORD], " & _
						"[SECURITY_QUESTION], " & _
						"[REMINDER], " & _
						"[ANSWER])" & _
						" values('" & request("PERSONIDnew") & "', " & _
						"'" & replace(request("USERNAMEnew"),"'","''") & "', " & _
						"'" & replace(request("PASSWORDnew"),"'","''") & "', " & _
						"'" & replace(request("SECURITYQUESTIONnew"),"'","''") & "', " & _
						"'" & replace(request("REMINDERnew"),"'","''") & "', " & _
						"'" & replace(request("ANSWERnew"),"'","''") & "')"

				cn.Execute(sql)
				session("message") = session("message") & "A new user login has been added.<br />"
			end if


		end if ''btnsave

	end if ''request.totalbytes

	CHOICE = request("choice")

	if CHOICE = "ALL" Then
		selectCHOICE ""
		printform CHOICE
	elseif CHOICE <> "" then
		selectCHOICE CHOICE
		printform CHOICE
	else
		'default to last names of 'A'
		selectCHOICE "A"
		'printform CHOICE
		'CHOICE ="A"
		printform "A"
	end if
%>
<!--#include virtual="/incs/fragContact.asp"-->

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub selectCHOICE(CHOICE)
%>
	<table bgcolor='#FFFFFF' align='center'>
	<tr><td>
		<font class='cfont8'><b>
		<ul>
		<li>Please select a person to add login information.</li>
		<li>User the filter dropdown to change the list of people displayed.</li>
		</ul>
		</b></font>
	</td></tr>
	</table>

	<br />
	<form name='filter'>
	<div align='center'>
		<font class='cfont10'>
		Last Initial:
			<select name='CHOICE' onChange="changePage();">
			<option value='A'>A</option>
			<option value='B'>B</option>
			<option value='C'>C</option>
			<option value='D'>D</option>
			<option value='E'>E</option>
			<option value='F'>F</option>
			<option value='G'>G</option>
			<option value='H'>H</option>
			<option value='I'>I</option>
			<option value='J'>J</option>
			<option value='K'>K</option>
			<option value='L'>L</option>
			<option value='M'>M</option>
			<option value='N'>N</option>
			<option value='O'>O</option>
			<option value='P'>P</option>
			<option value='Q'>Q</option>
			<option value='R'>R</option>
			<option value='S'>S</option>
			<option value='T'>T</option>
			<option value='U'>U</option>
			<option value='V'>V</option>
			<option value='W'>W</option>
			<option value='X'>X</option>
			<option value='Y'>Y</option>
			<option value='Z'>Z</option>
			<option value=''>All</option>
			</select>
		</font>
	</div>
	</form>

	<script language='Javascript'>
	<!--

		function changePage()
		{
			choice = filter.CHOICE.options[filter.CHOICE.selectedIndex].value;
			//alert(choice);
			//if(choice != "")
			//{
				filter.submit();
			//}
		}
	//-->
	</script>

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub printform(CHOICE)
	''list of people with no login information in USER_LOGIN_TBL
	sql = "SELECT PERSON_ID, "_
		& "LAST_NAME + ', ' + FIRST_NAME as 'NAME' "_
		& "FROM db_accessadmin.[PERSON_TBL] "_
		& "WHERE LAST_NAME LIKE ('" & CHOICE & "%') "_
		& "AND PERSON_ID NOT IN (SELECT PERSON_ID FROM USER_LOGIN_TBL) "_
		& "ORDER BY UPPER(LAST_NAME),UPPER(FIRST_NAME)"
	'Response.Write(sql)
	'Response.End
	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsAvailablePersonData = rs.GetRows
		rsAvailablePersonRows = UBound(rsAvailablePersonData,2)
	else
		rsAvailablePersonRows = -1
	end if

	sql = "SELECT l.PERSON_ID, "_
		& "p.LAST_NAME + ', ' + p.FIRST_NAME as 'NAME', "_
		& "[USERNAME], "_
		& "[PASSWORD], "_
		& "[SECURITY_QUESTION], "_
		& "[REMINDER], "_
		& "[ANSWER] "_
		& "FROM USER_LOGIN_TBL l "_
		& "LEFT JOIN db_accessadmin.[PERSON_TBL] p ON l.PERSON_ID = p.PERSON_ID "_
		& "WHERE LAST_NAME LIKE ('" & CHOICE & "%') "_
		& "ORDER BY UPPER(LAST_NAME),UPPER(FIRST_NAME)"
	''rw(sql)
	''Response.End
	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsPersonData = rs.GetRows
		rsPersonRows = UBound(rsPersonData,2)
	else
		rsPersonRows = -1
	end if

	if session("message") <> "" then
		response.write "<div class=""cfontSuccess10"" align='center'>" & session("message") & "</div>"
		session("message") = ""
	end if

%>

<form method="post" name='MODIFY'>

<table width='500' bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>
<tr bgcolor='#000066'>
<th><font class='cfontWhite10'>Remove?</font></th>
<th><font class='cfontWhite10'>Person</font></th>
<th><font class='cfontWhite10'>User Name</font></th>
<th><font class='cfontWhite10'>Password</font></th>
<th><font class='cfontWhite10'>Security Question</font></th>
<th><font class='cfontWhite10'>Reminder</font></th>
<th><font class='cfontWhite10'>Answer</font></th>
</tr>

<tr>
<th colspan="7"><font class='cfont10'>New Login Information</font></th>
</tr>

<tr BGCOLOR = "#FFFFFF">
<td>&nbsp;</td>
<td><%makedd "PERSONIDnew" %></td>
<td><input type="text" maxlength="20" name="USERNAMEnew" id="USERNAMEnew" /></td>
<td><input type="text" maxlength="12" name="PASSWORDnew" id="PASSWORDnew" /></td>
<td><input type="text" maxlength="100" name="SECURITYQUESTIONnew" id="SECURITYQUESTIONnew" /></td>
<td><input type="text" maxlength="30" name="REMINDERnew" id="REMINDERnew" /></td>
<td><input type="text" maxlength="20" name="ANSWERnew" id="ANSWERnew" /></td>
</tr>

<tr>
<td colspan="7"><input type="submit" name="btnsave" value="Submit" class="btn" /></td>
</tr>

<%
	For i = 0 to rsPersonRows

		ODD_ROW = NOT ODD_ROW

		If ODD_ROW Then
			BGCOLOR = "#FFFFFF"
		Else
			BGCOLOR = "#F0F0F0"
		End If
%>
		<tr bgcolor='<%=BGCOLOR%>'>
		<td align='center'><input type='checkbox' name='REMOVE' value='<%=rsPersonData(0,i)%>' /></td>
		<td><font class='cfont10'><%=rsPersonData(1,i)%></font></td>
		<td><input type="text" maxlength='20' name="USERNAME_<%=rsPersonData(0,i)%>" id="USERNAME_<%=rsPersonData(0,i)%>" value="<%=rsPersonData(2,i)%>" /></td>
		<td><input type="text" maxlength='12' name="PASSWORD_<%=rsPersonData(0,i)%>" id="PASSWORD_<%=rsPersonData(0,i)%>" value="<%=rsPersonData(3,i)%>" /></td>
		<td><input type="text" maxlength='100' name="SECURITYQUESTION_<%=rsPersonData(0,i)%>" id="SECURITYQUESTION_<%=rsPersonData(0,i)%>" value="<%=rsPersonData(4,i)%>" /></td>
		<td><input type="text" maxlength='30' name="REMINDER_<%=rsPersonData(0,i)%>" id="REMINDER_<%=rsPersonData(0,i)%>" value="<%=rsPersonData(5,i)%>" /></td>
		<td><input type="text" maxlength='20' name="ANSWER_<%=rsPersonData(0,i)%>" id="ANSWER_<%=rsPersonData(0,i)%>" value="<%=rsPersonData(6,i)%>" /></td>
		</tr>
<%
	Next
%>

</table>

</form>

<%

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub makedd(name)
%>
<select name="<%=name%>" id="<%=name%>">
<option value="">- Select A Person -</option>
<%
	For j = 0 to rsAvailablePersonRows
%>
		<option value="<%=rsAvailablePersonData(0,j)%>"><%=rsAvailablePersonData(1,j)%></option>
<%
	Next

%>
</select>

<%
end sub

end class


set ag = new userlogin

%>


