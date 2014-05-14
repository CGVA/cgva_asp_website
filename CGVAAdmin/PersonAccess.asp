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
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Person Access Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Person Access Administration</u></b>
	</font>
	</div>

	<br />

<%

class personaccess

sub class_initialize()

	PERSON_ID = request("PERSON_ID")

	if request.totalbytes=0 then

	else

		if request("btnsave") <> "" then

			''rw(i & "<br />")
			''rw(request("ACCESS") & "<br />")
			arrid = split(request("ACCESS"), ", ")
			''rw(arrid(1) & "<br />")
			''rw(arrid(2) & "<br />")

			''delete all for the selected person and then restart
			sql = "DELETE FROM PERSON_ACCESS_TBL WHERE PERSON_ID = '" & PERSON_ID & "'"
			''rw(sql & "<br />")
			cn.Execute(sql)

			For i = 0 to UBound(arrid)
				sql = 	"insert into PERSON_ACCESS_TBL(" & _
						"PERSON_ID,ROLE_TYPE) VALUES(" & _
						"'" & PERSON_ID & "', " & _
						"'" & arrid(i) & "')"
					''response.write sql & "<br />"
				cn.Execute(sql)
			Next

			session("message") = session("message") & "Access for the selected person has been updated successfully.<br />"

		end if

	end if

	if PERSON_ID <> "" then
		selectPerson PERSON_ID
		printform PERSON_ID
	else
		selectPerson ""
	end if

%>
<!--#include virtual="/incs/fragContact.asp"-->

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub selectPerson(PERSON_ID)
	sql = "SELECT PERSON_ID, LAST_NAME + ', ' + FIRST_NAME "_
		& "FROM db_accessadmin.PERSON_TBL "_
		& "WHERE LOGICAL_DELETE_IND = 'N' "_
		& "ORDER BY UPPER(LAST_NAME), UPPER(FIRST_NAME)"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsPersonData = rs.GetRows
		rsPersonRows = UBound(rsPersonData,2)
	else
		rw("Error:Missing Active People.")
		Response.End
	end if
%>
	<table align='center'>
	<tr><td>
		<font class='cfont8'><b>
		<ul>
		<li>Instructions</li>
		</b></font>
		</ul>
	</td></tr>
	</table>

	<br />
	<form name='personChoice'>
	<div align='center'>
		<font class='cfont10'>
		Name:
		<select name='PERSON_ID' onChange="changePage();">
		<option value=''>-select-</option>
		<%
			For i = 0 to rsPersonRows
				rw("<option value='" & rsPersonData(0,i) & "'")
					If CStr(PERSON_ID) = CStr(rsPersonData(0,i)) then
					rw(" selected")
				End If
					rw(">" & rsPersonData(1,i) & "</option>")
			Next
		%>
		</select>
		</font>
	</div>
	</form>

	<script language='Javascript'>
	<!--

		function changePage()
		{
			choice = personChoice.PERSON_ID.options[personChoice.PERSON_ID.selectedIndex].value;
			//alert(choice);
			if(choice != "")
			{
				personChoice.submit();
			}
		}
	//-->
	</script>

<%
end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub printform(PERSON_ID)
	sql = "SELECT * FROM ROLE_TYPE_TBL ORDER BY UPPER(ROLE_TYPE_DESC)"
	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsRTData = rs.GetRows
		rsRTRows = UBound(rsRTData,2)
	else
		rsRTRows = -1
	end if

	sql = "SELECT ROLE_TYPE FROM PERSON_ACCESS_TBL WHERE PERSON_ID = '" & PERSON_ID & "'"

	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsPAData = rs.GetRows
		rsPARows = UBound(rsPAData,2)

		For i = 0 to rsPARows
			ACCESS_SELECTED = ACCESS_SELECTED & rsPAData(0,i) & ","
		Next
	else
		ACCESS_SELECTED = ""
	end if

	if session("message") <> "" then
		response.write "<div class=""cfontSuccess10"" align='center'>" & session("message") & "</div>"
		session("message") = ""
	end if

%>

<form method="post" name='MODIFY'>

<table width='500' cellspacing='1' align='center' cellpadding='3'>
<tr>
<th valign='top'><font class='cfont10'><nobr>Access Level(s)</nobr></font></th>
<th valign='top'><%makedd rsRTData,rsRTRows, "ACCESS", ACCESS_SELECTED %></th>
<td valign='top'><input type="submit" name="btnsave" value="Submit" class="btn" /></td>
</tr>
</table>

</form>

<%

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sub makedd(rs, rsRows, name, selected)
%>
<select name="<%=name%>" id="<%=name%>" multiple size='5'>
<%
	For i = 0 to rsRows

%>

		<option<%If InStr(selected, rs(0,i)) > 0 then rw(" selected") End If%> value="<%=rs(0,i)%>"><%=rs(0,i)%> - <%=rs(1,i)%></option>

<%
	Next

%>
</select>

<%
end sub

end class


set pa = new personaccess

%>


