<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Password.aspx.vb" Inherits="Password" Classname="MyEvents" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>MyCGVA - My Password Information</title>
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
    'JPC TEST'
    Response.WriteFile("~/incs/fragMyCGVA_DEVNavigation.asp")
%>
</td>
<td>&nbsp;</td>
<td valign='top' width='100%'>

	<table width='100%' bgcolor="#FFFFFF" border="0">

	<tr bgcolor="#FF3300">
	<td height="20" class="cfont16">MyCGVA - My Password Information</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

        <form id="INSERT" runat="server">

		<div align='center'>
		<font class='cfont10'><b>Listed below is your current password information. Please update any fields, and click 'Submit New Information'.<br />
		Any new password information will be applied to your next visit to MyCGVA.</b></font>
		<p />
		</div>

        <table cellspacing='0' align='center' cellpadding='3'>

        <tr>
        <td colspan='2' align='center'>
        <font class="cfontError10">
        <asp:Label ID="messageLabel" runat="server" Text=""></asp:Label>
        </font>
        </td>
        </tr>

        <tr>
        <td align='right' valign='middle' class="style1"><font size='2' face='arial'>User Name</font></td>
        <td class="style1"><asp:TextBox ID="USERNAME" size='25' maxlength='25' runat="server" enabled="false"></asp:TextBox>
        </td>
        </tr>

        <tr>
        <td align="right">
        <font class="cfont10">Password</font></font><font class="cfontError10">*</font></td>
        <td>
        <font class="cfont10"><asp:TextBox ID="PASSWORD1" size='20' maxlength='20' 
        runat="server"></asp:TextBox>
        (between 
        6 and 20 characters)</td>
        </tr>
        <tr bgcolor="#ffffff">
        <td align="right">
        <font class="cfont10">Re-Type password:<font class="cfontError10">*</font></td>
        <td>
        <asp:TextBox ID="PASSWORD2" size='20' maxlength='20' 
        runat="server"></asp:TextBox></td>
        </tr>
        <tr bgcolor="#eeeeee">
        <td align="right">
        <font class="cfont10">Security Question:<font class="cfontError10">*</font></td>
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
        <font class="cfont10">Security Answer:<font class="cfontError10">*</font></td>
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
</table>
<script type="text/javascript" language='Javascript'>
<!--
	INSERT.PASSWORD1.focus();
//-->
</script>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



