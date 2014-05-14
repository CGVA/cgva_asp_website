<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="Login" enableEventValidation="false" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>MyCGVA - Login</title>

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
	<td height="20" class="cfont16">MyCGVA - Login</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

     <tr>
     <td valign='top'>
        <form id="loginForm" runat="server">
	    <table align='center' cellpadding='3'>
	    <tr>
	    <td><font class='cfont10'>User Name</font></td>
	    <td>
            <asp:TextBox name="USERNAME" ID="USERNAME" maxlength="50" runat="server"></asp:TextBox>
                            </td>
	    </tr>
	    <tr>
	    <td><font class='cfont10'>Password</font></td>
	    <td>
            <asp:TextBox textmode="password" ID="PASSWORD" maxlength="20" runat="server"></asp:TextBox>
                            </td>
	    </tr>
	    <tr>
	    <td>&nbsp;</td>
	    <td><font class='cfont10'><asp:Button ID="submitButton" runat="server" 
                Text="Continue" />
            </font></td>
	    </tr>

	    <tr>
	    <td>&nbsp;</td>
	    <td><font class='cfont10'>Not a member yet? <a href='NewRegistration.aspx'>Click here</a></font></td>
	    </tr>

	    <tr>
	    <td>&nbsp;</td>
	    <td><font class='cfont10'>Forgot User Name/Password? <a href='LoginHelp.aspx'>Click here</a></font></td>
	    </tr>

	    <tr>
	    <td colspan="2"><font class='cfontError10'><asp:Label ID="ErrorLabel" runat="server" Text=""></asp:Label></font></td>
	    </tr>

	    <tr>
	    <td colspan="2" height="200">&nbsp;</td>
	    </tr>
	    	    </table>
        </form>
    </td></tr>

    </table>
    
</td></tr>
</table>

<script type="text/javascript" language="javascript">
<!--
    //alert("here");
    loginForm.USERNAME.focus();
-->
</script>
</body>
</html>

