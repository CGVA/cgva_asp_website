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

    Dim rsTeamData, rsTeamRows
%>

<!-- #include virtual="/incs/fragHeader.asp" -->
<title>CGVA - Last Dig Payment Administration</title>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td valign='top'>

	<div align='center'>
	<font class='cfont12'>
	<b><u>CGVA - Last Dig Payment Administration</u></b>
	</font>
	</div>

	<br />

<%

class teamPayment


sub class_initialize()
	if request.totalbytes=0 then

	else

		if request("btnsave") <> "" then

			dim i
                    
                    For Each i In Request.form

                        arrid = Split(i, "_")

                        ''Response.Write( arrid(0) & "<br />")
                        If arrid(0) = "TRANSACTIONID" Then

                            if Request("TRANSACTIONDATE_" & arrid(1)) = "" then vDate = "NULL" else vDate = "'" & Request("TRANSACTIONDATE_" & arrid(1)) & "'" end if
                            if Request("PAYMENTVERIFIED_" & arrid(1)) = "" then vDate2 = "NULL" else vDate2 = "'" & Request("PAYMENTVERIFIED_" & arrid(1)) & "'" end if

                            sql = "update LASTDIG_TEAM_TBL " & _
                              "set [TEAM_NAME] = '" & Replace(Request("TEAMNAME_" & arrid(1)), "'", "''") & "', " & _
                              "[DIVISION] = '" & Replace(Request("DIVISION_" & arrid(1)), "'", "''") & "', " & _
                              "ID_DTG = " & vDate & ", " & _
                              "FEE_PAID = '" & Replace(Request("PAYMENTRECEIVED_" & arrid(1)), "'", "''") & "', " & _
                              "ACTIVE = '" & Replace(Request("ACTIVE_" & arrid(1)), "'", "''") & "', " & _
                              "RECONCILE_DTG = " & vDate2 & " " & _
                              "where ID = '" & arrid(1) & "'"
                            ''response.write sql & "<br />"
                            cn.Execute(sql)

                        End If

                    Next 'i

                    Session("message") = Session("message") & "Last Dig Team Registration information has been updated successfully.<br />"

                    If Request("TRANSACTIONDATENew") <> "" Then

                        if Request("PAYMENTVERIFIEDNew") = "" then vDate = "NULL" else vDate = "'" & Request("PAYMENTVERIFIEDNew") & "'" end if

                        sql = "insert into LASTDIG_TEAM_TBL" & _
                          "([ID_DTG], " & _
                          "[TEAM_NAME], " & _
                          "[DIVISION], " & _
                          "[FEE_PAID], " & _
                          "[RECONCILE_DTG], " & _
                          "[ACTIVE])" & _
                          " values(getdate(), " & _
                          "'" & Replace(Request("TEAMNAMENew"), "'", "''") & "', " & _
                          "'" & Replace(Request("DIVISIONNew"), "'", "''") & "', " & _
                          "'" & Replace(Request("PAYMENTRECEIVEDNew"), "'", "''") & "', " & _
                          vDate & ", " & _
                        "'" & Replace(Request("ACTIVENew"), "'", "''") & "')"
                        
                        'response.write sql & "<br />"
                        cn.Execute(sql)
                        Session("message") = Session("message") & "A new team registration has been added.<br />"
                    End If

                       'Response.End

		end if ''btnsave

	end if ''request.totalbytes

	CHOICE = request("CHOICE")

	if CHOICE = "" Then
		Session("CHOICE") = "Y"
		selectCHOICE "Y"
		printform "Y"
	elseif CHOICE <> "" then
		Session("CHOICE") = CHOICE
		selectCHOICE CHOICE
		printform CHOICE
	else
       Response.Write("HERE")
        'default to last names of Active ('Y')
        Session("CHOICE") = "Y" 
        selectCHOICE "Y" 
		printform CHOICE
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
		<li>Please enter team registration information and/or modify an existing registration 
            in the fields provided.</li>
		<li>User the filter dropdown to view active/inactive.</li>
		</ul>
		</b></font>
	</td></tr>
	</table>

	<br />
	<form name='filter'>
	<div valign="middle" align='center'>
		<font class='cfont10'>
		Active/Inactive:
			<select name='CHOICE' onChange="changePage();">
			<option value='Y'<%If Session("CHOICE")="Y" or Session("CHOICE")="" Then rw(" selected") End If %>>Active</option>
			<option value='N'<%If Session("CHOICE")="N" Then rw(" selected") End If %>>Inactive</option>
			<option value='ALL'<%If Session("CHOICE")="ALL" Then rw(" selected") End If %>>All</option>
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
    ''list of inactive/active  with no login information in USER_LOGIN_TBL
    sql = "SELECT ID, " _
     & "[ID_DTG], " _
     & "TEAM_NAME, " _
     & "DIVISION, " _
     & "[FEE_PAID], " _
     & "[RECONCILE_DTG], " _
     & "[ACTIVE] " _
     & "FROM LASTDIG_TEAM_TBL " _
     & "WHERE ACTIVE LIKE ('%" & Replace(CHOICE,"ALL","") & "%') " _
     & "ORDER BY DIVISION,UPPER(TEAM_NAME)"
	''rw(sql)
	''Response.End
	set rs = cn.Execute(sql)

	if not rs.EOF then
        rsTeamData = rs.GetRows
        rsTeamRows = UBound(rsTeamData, 2)
	else
        rsTeamRows = -1
	end if

	if session("message") <> "" then
		response.write "<div class=""cfontSuccess10"" align='center'>" & session("message") & "</div>"
		session("message") = ""
	end if

%>

<form method="post" name='MODIFY'>

<table width='500' bgcolor='#9999FF' cellspacing='1' align='center' cellpadding='3'>
<tr bgcolor='#000066'>
<th><font class='cfontWhite10'>Transaction ID</font></th>
<th><font class='cfontWhite10'>Transaction Date</font></th>
<th><font class='cfontWhite10'>Team Name</font></th>
<th><font class='cfontWhite10'>Division</font></th>
<th><font class='cfontWhite10'>Fee Paid</font></th>
<th><font class='cfontWhite10'>Payment Verified</font></th>
<th><font class='cfontWhite10'>Active</font></th>
</tr>

<tr>
<th colspan="7"><font class='cfont10'>New Information</font></th>
</tr>

<tr BGCOLOR = "#FFFFFF">
<td align="center"><input type="text" size='2' maxlength="4" name="TRANSACTIONIDNew" id="TRANSACTIONIDNew" disabled/></td>
<td align="center"><input type="text" size='17' maxlength="30" name="TRANSACTIONDATENew" id="TRANSACTIONDATENew" /></td>
<td align="center"><input type="text" size='35' maxlength="50" name="TEAMNAMENew" id="TEAMNAMENew" /></td>
<td align="center">
<select name="DIVISIONNew" id="DIVISIONNew">
<option value='B'>B</option>
<option value="BB">BB</option>
<option value="Modified A">Modified A</option>
</select>
</td>
<td align="center"><input type="text" size='8' maxlength="10" name="PAYMENTRECEIVEDNew" id="PAYMENTRECEIVEDNew" /></td>
<td align="center"><input type="text" size='17' maxlength="20" name="PAYMENTVERIFIEDNew" id="PAYMENTVERIFIEDNew" /></td>
<td align="center">
<select name="ACTIVENew" id="ACTIVENew">
<option value='Y'>Yes</option>
<option value="N">No</option>
</select>
</td>
</tr>

<tr>
<td colspan="7"><input type="submit" name="btnsave" value="Submit" class="btn" /></td>
</tr>

<%
    For i = 0 To rsTeamRows

        ODD_ROW = Not ODD_ROW

        If ODD_ROW Then
            BGCOLOR = "#FFFFFF"
        Else
            BGCOLOR = "#F0F0F0"
        End If
%>
		<tr bgcolor='<%=BGCOLOR%>'>
		<td align="center">
		    <input type="hidden" name="TRANSACTIONID_<%=rsTeamData(0,i)%>" id="TRANSACTIONID_<%=rsTeamData(0,i)%>" value="<%=rsTeamData(0,i)%>" />
		    <%=rsTeamData(0,i)%>
		</td>
		<td align="center"><input type="text" size='17' maxlength='30' name="TRANSACTIONDATE_<%=rsTeamData(0,i)%>" id="TRANSACTIONDATE_<%=rsTeamData(0,i)%>" value="<%=rsTeamData(1,i)%>" /></td>
		<td align="center"><input type="text" size='35' maxlength='50' name="TEAMNAME_<%=rsTeamData(0,i)%>" id="TEAMNAME_<%=rsTeamData(0,i)%>" value="<%=rsTeamData(2,i)%>" /></td>
		<td align="center">
            <select name="DIVISION_<%=rsTeamData(0,i)%>" id="DIVISION_<%=rsTeamData(0,i)%>">
            <option value='B'<%If rsTeamData(3,i) ="B" then rw(" selected") end if%>>B</option>
            <option value="BB"<%If rsTeamData(3,i) ="BB" then rw(" selected") end if%>>BB</option>
            <option value="Modified A"<%If rsTeamData(3,i) ="Modified A" then rw(" selected") end if%>>Modified A</option>
            </select>
        </td>
		<td align="center"><input type="text" size='5' maxlength='10' name="PAYMENTRECEIVED_<%=rsTeamData(0,i)%>" id="PAYMENTRECEIVED_<%=rsTeamData(0,i)%>" value="<%=rsTeamData(4,i)%>" /></td>
		<td align="center"><input type="text" size='17' maxlength='20' name="PAYMENTVERIFIED_<%=rsTeamData(0,i)%>" id="PAYMENTVERIFIED_<%=rsTeamData(0,i)%>" value="<%=rsTeamData(5,i)%>" /></td>
		<td align="center">
            <select name="ACTIVE_<%=rsTeamData(0,i)%>" id="ACTIVE_<%=rsTeamData(0,i)%>">
            <option value='Y'<%If rsTeamData(6,i) ="Y" then rw(" selected") end if%>>Yes</option>
            <option value="N"<%If rsTeamData(6,i) ="N" then rw(" selected") end if%>>No</option>
            </select>
		</td>
		</tr>
<%
	Next
%>

</table>

</form>

<%

end sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

end class


set ag = new teamPayment

%>



