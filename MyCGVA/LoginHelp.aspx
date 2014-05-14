<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LoginHelp.aspx.vb" Inherits="LoginHelp" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - Login Assistance</title>

<%		   
    Response.WriteFile("~/incs/fragHeader.asp")
%>

</head>
<%		   
    Response.WriteFile("~/incs/Header.asp")
%>
<table bgcolor='#000066' width='780' align='center' border='0' cellpadding='0' cellspacing='0'>
<tr>
<td colspan='3'>
<%		   
    Response.WriteFile("~/incs/fragSiteHeaderGraphics.asp")
%>
</td>
</tr>

<tr>
<td valign='top' width='175'>
<%		   
    If Session("PERSON_ID") <> "" And Session("TEMP_PASSWORD") = "" Then
        Response.WriteFile("~/incs/fragMyCGVANavigation.asp")
    Else
%>
<table width='175' align='left' border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">

<tr bgcolor="#000066">
<td align='center'><a class='menuBlack' href="http://www.cgva.org"><font class='cfontWhite12'><b>Return to CGVA</b></font></a></td>
</tr>
</table>
<%
    End If
%>
</td>
<td>&nbsp;</td>
<td valign='top' width='100%'>
	<table width='100%' bgcolor="#FFFFFF" border="0">

	<tr bgcolor="#FF3300">
	<td height="20" class="cfont16">MyCGVA - Login Assistance</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

 <tr>
 <td valign='top'>
    <form id="loginForm" runat="server">
	<table align='center' cellpadding='3'>
	
	<tr><td><font class='cfont10'><b>Forgot your user name/password? You can have it 
        sent to you from here!<br />
        </b>&nbsp;<br />
        Please enter your email address below and click 'Request Login Information'. 
	    <br />
        Your user name and a new temporary password will be sent to this email address shortly.<br />
	<p />
	Email Address: 
        <asp:TextBox ID="EMAIL_ADDRESS" size="40" maxlength='50' runat="server"></asp:TextBox>
        </font>&nbsp;
        <asp:Button ID="submitButton" runat="server" Text="Request Login Information" />
	</td>
	</tr>
    <tr>
    <td>
         <font class="cfontSuccess10"><asp:Label ID="successLabel" runat="server" Text=""></asp:Label></font>
    </td>
    </tr>

    <tr>
    <td height="500">
         &nbsp;
    </td>
    </tr>

	</table>
    </form>
</td></tr>
</table>
<script type="text/javascript" language='Javascript'>
<!--
	loginForm.EMAIL_ADDRESS.focus();
//-->
</script>
</body>
</html>


