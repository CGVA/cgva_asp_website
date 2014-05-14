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
<title>CGVA - Role Type Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Role Type Administration</u></b>
	</font>
	</div>

	<br />

<%
	if request.totalbytes=0 then

	else

		if request("btnsave") <> "" then

			dim i

			''deal with deletes
			for each i in request.form

				if i = "REMOVE" then
					arrRemove = split(request("REMOVE"), ", ")

					for j = 0 to UBound(arrRemove)
						arrRemoveParts = split(arrRemove(j),"_")

						SQL = 	"DELETE FROM ROLE_TYPE_TBL " & _
								"WHERE ROLE_TYPE='" & Replace(arrRemoveParts(0),"'","''") & "' AND ROLE_TYPE_DESC='" & Replace(arrRemoveParts(1),"'","''") & "'"
						cn.Execute(SQL)

					next

				end if

			next

			for each i in request.form
				''rw(i & "<br />")
				arrid = split(i, "_")
				''rw(arrid(0) & "<br />")
				''rw(arrid(1) & "<br />")

				if arrid(0) = "ROLETYPE" then

					sql = 	"update ROLE_TYPE_TBL " & _
							"set [ROLE_TYPE] = '" & replace(request("ROLETYPE_" & arrid(1) & "_" & arrid(2)),"'","''") & "', " & _
							"[ROLE_TYPE_DESC] = '" & replace(request("ROLETYPEDESC_" & arrid(1) & "_" & arrid(2)),"'","''") & "' " & _
							"WHERE ROLE_TYPE = '" & arrid(1) & "' " & _
							"AND ROLE_TYPE_DESC = '" & arrid(2) & "'"
					''response.write sql & "<br />"
					cn.Execute(sql)

				end if

			next 'i

			session("message") = session("message") & "Role Types have been updated successfully.<br />"

			if request("ROLETYPEnew") <> "" and request("ROLETYPEDESCnew") <> "" then

				sql = "select * from ROLE_TYPE_TBL rtt "_
					& "WHERE ROLE_TYPE = '" & replace(request("ROLETYPEnew"),"'","''") & "' "_
					& "AND ROLE_TYPE_DESC = '" & replace(request("ROLETYPEDESCnew"),"'","''") & "'"
				set rs = cn.Execute(sql)

				if not rs.EOF then
					session("message") = session("message") & "<font class='cfontError10'>The new game number entered already exists for the selected match. Please check your information, and try entering the game information again.</font>"
				else
					sql = 	"insert into ROLE_TYPE_TBL" & _
							"(ROLE_TYPE, " & _
							"ROLE_TYPE_DESC)" & _
							" values('" & replace(request("ROLETYPEnew"),"'","''") & "', " & _
							"'" & replace(request("ROLETYPEDESCnew"),"'","''") & "')"
					''response.write sql & "<br />"
					cn.Execute(sql)
					session("message") = session("message") & "A new role type has been added.<br />"
				end if

			end if

		end if ''btnsave

	end if ''request.totalbytes

%>

<table align='center'>
<tr><td>
	<font class='cfont8'><b>
	<ul>
	<li>Instructions</li>
	</ul>
	</b></font>
</td></tr>
</table>

<br />

<%

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = "SELECT * FROM ROLE_TYPE_TBL ORDER BY UPPER(ROLE_TYPE_DESC)"
	set rs = cn.Execute(sql)

	if not rs.EOF then
		rsData = rs.GetRows
		rsRows = UBound(rsData,2)
	else
		rsRows = -1
	end if

	if session("message") <> "" then
		response.write "<div class=""cfontSuccess10"" align='center'>" & session("message") & "</div>"
		session("message") = ""
	end if

%>

<form method="post" name='MODIFY'>

<table bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>
<tr bgcolor='#000066'>
<th><font class='cfontWhite10'>Remove?</font></th>
<th><font class='cfontWhite10'>Role Type</font></th>
<th><font class='cfontWhite10'>Description</font></th>
</tr>

<tr>
<th colspan="3"><font class='cfont10'>New Role Type</font></th>
</tr>

<tr bgcolor='#FFFFFF'>
<td>&nbsp;</td>
<td><input type="text" maxlength="10" style="width:200px;" name="ROLETYPEnew" id="ROLETYPEnew" /></td>
<td><input type="text" maxlength="60" style="width:400px;" name="ROLETYPEDESCnew" id="ROLETYPEDESCnew" /></td>
</tr>

<tr>
<td colspan="9"><input type="submit" name="btnsave" value="Submit" class="btn" /></td>
</tr>

<%
	For i = 0 to rsRows

		ODD_ROW = NOT ODD_ROW

		If ODD_ROW Then
			BGCOLOR = "#FFFFFF"
		Else
			BGCOLOR = "#F0F0F0"
		End If
%>
		<tr bgcolor='<%=BGCOLOR%>'>
		<td align='center'><input type='checkbox' name='REMOVE' value='<%=rsData(0,i)%>_<%=rsData(1,i)%>' /></td>
		<td><input type="hidden" name="ROLETYPE_<%=rsData(0,i)%>_<%=rsData(1,i)%>" id="ROLETYPE_<%=rsData(0,i)%>_<%=rsData(1,i)%>" value="<%=rsData(0,i)%>" /><font class='cfont10'><%=rsData(0,i)%></font></td>
		<td><input type="text" maxlength='60' style="width:400px;" name="ROLETYPEDESC_<%=rsData(0,i)%>_<%=rsData(1,i)%>" id="ROLETYPEDESC_<%=rsData(0,i)%>_<%=rsData(1,i)%>" value="<%=rsData(1,i)%>" /></td>
		</tr>
<%
	Next
%>

</table>

</form>

<!--#include virtual="/incs/fragContact.asp"-->

