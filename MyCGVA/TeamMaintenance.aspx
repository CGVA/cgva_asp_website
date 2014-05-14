<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TeamMaintenance.aspx.vb" Inherits="TeamMaintenance" enableEventValidation="false" %>

<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>MyCGVA - My Teams</title>
<%		   
    Response.WriteFile("/incs/fragHeader.asp")
%>

</head>
<%		   
    Response.WriteFile("/incs/Header.asp")
%>
<table bgcolor='#000066' width='780' align='center' border='0' cellpadding='0' cellspacing='0'>
<tr>
<td colspan='3'>
<%		   
    Response.WriteFile("/incs/fragSiteHeaderGraphics.asp")
%>
</td>
</tr>

<tr>
<td valign='top' width='175'>
<%		   
    'JPC TEST'
    Response.WriteFile("/incs/fragMyCGVANavigation.asp")
%>
</td>
<td>&nbsp;</td>
<td valign='top' width='100%'>

	<table width='100%' bgcolor="#FFFFFF" border="0">

	<tr bgcolor="#FF3300">
	<td height="20" class="cfont16">MyCGVA - My Teams - Maintenance</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

		<div align='center'>
		<font class='cfont10'><b><asp:label runat="server" ID="messageLabel"></asp:label></b></font>
        <p />
		<font class='cfont10'><b><asp:label runat="server" ID="successLabel"></asp:label></b></font>
		</div>

		<form id="Form1" name='INSERT' runat="server">
		<table align='center'>
		<tr>
		<td>

		    <asp:Table ID="teamTable" align='center' cellpadding='3' runat="server">
                    
            <asp:tablerow bgcolor="#C0C0C0">
            <asp:tablecell align="center">&nbsp;</asp:tablecell>
            <asp:tablecell align="center"><font class='cfont10'><b>Player Name</b></font></asp:tablecell>
            </asp:tablerow>
            
           </asp:table>
        </td>
        </tr>
        
        <tr>
        <td align="right">
            <asp:Button id="submitButton" text="Save My Team" runat="server" />
        </td>
        </tr>
        </table>

        <asp:hiddenfield ID="ID" runat="server"></asp:hiddenfield>
        <asp:hiddenfield ID="EVENT_CD" runat="server"></asp:hiddenfield>
        </form>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



