<%@ Page Language="VB" AutoEventWireup="false" CodeFile="VerifyPassword.aspx.vb" Inherits="VerifyPassword" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - Reset Temporary Password</title>

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
	<td height="20" class="cfont16">MyCGVA - Reset Temporary Password</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

</td>
</tr>
 
 <tr>
 <td valign='top'>
    <form id="INSERT" runat="server">
	<table align='center' cellpadding='3'>
	
	<tr><td colspan='2'>
	<font class='cfont10'>Welcome new MyCGVA member!<p />
	Please enter the following password information to continue. 
    <br />
    Your password should be between 6 and 20 characters in length, and should not be identical to your User 
        Name.
        <br />
    Once you have finished with this step, you will be able to access all areas of MyCGVA.</font></td></tr>
	
		<table cellspacing='0' align='center' cellpadding='3'>

        <tr>
        <td colspan='2' align='center'>
        <font class="cfontError10">
        <asp:Label ID="messageLabel" runat="server" Text=""></asp:Label>
        </font>
        </td>
        </tr>
		
       <tr>
        <td  align="right" bgcolor="#ffffff">
        <font class="cfont12"><b>Create Password</b></font><font class="cfontError10">*</font></td>
        <td align="right" bgcolor="#ffffff" valign="top">
        &nbsp;</td>
        </tr>

				<tr>
		<td align='right' valign='middle' class="style1"><font class="cfont10">User Name</font></td>
		<td class="style1"><asp:TextBox ID="USERNAME" size='25' maxlength='25' runat="server" enabled="false"></asp:TextBox>
            </td>
		</tr>
		
		<tr>
        <td align="right">
        <font class="cfont10">Password</font></font><font class="cfontError10">*</font></td>
        <td>
            <font class="cfont10"><asp:TextBox textmode="password" ID="PASSWORD1" size='20' maxlength='20' 
                runat="server"></asp:TextBox>
                        (between 
        6 and 20 characters)</td>
        </tr>
        <tr bgcolor="#ffffff">
        <td align="right">
        <font class="cfont10">Re-Type password<font class="cfontError10">*</font></td>
        <td>
        <asp:TextBox textmode="password" ID="PASSWORD2" size='20' maxlength='20' 
                runat="server"></asp:TextBox></td>
        </tr>
        <tr bgcolor="#eeeeee">
        <td align="right">
        <font class="cfont10">Security Question<font class="cfontError10">*</font></td>
        <td>
        &nbsp;<asp:DropDownList ID="SECURITY_QUESTION" runat="server">
                <asp:ListItem Value="0">-- please make a selection --</asp:ListItem>
                <asp:ListItem Value="1">What is the name of your first pet?</asp:ListItem>
                <asp:ListItem Value="2">Your mother’s maiden name?</asp:ListItem>
                <asp:ListItem Value="3">Your father’s middle name?</asp:ListItem>
                <asp:ListItem Value="4">The street you grew up on?</asp:ListItem>
                <asp:ListItem Value="5">What is your favorite sports team?</asp:ListItem>
            </asp:DropDownList>
        
        </td>
        </tr>
        <tr bgcolor="#ffffff">
        <td align="right">
        <font class="cfont10">Security Answer<font class="cfontError10">*</font></td>
        <td>
        <asp:TextBox ID="SECURITY_ANSWER" size='30' maxlength='50' 
                runat="server"></asp:TextBox>
                        </td>
        </tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'>
            <asp:Button ID="submitButton" runat="server" Text="Submit Password Information" />
            </font></td>
		</tr>


		
		</table>

		        </form>
</td></tr>
<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>
<script type="text/javascript" language='Javascript'>
<!--
	INSERT.PASSWORD1.focus();
//-->
</script>
</body>
</html>


