<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyTeams.aspx.vb" Inherits="MyTeams" Classname="MyTeams" enableEventValidation="false" %>
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
	<td height="20" class="cfont16">MyCGVA - My Teams</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

		<div align='center'>
		<font class='cfont10'>*If you are a captain looking for team members, or a player looking for a team, click <a href="http://www.cgva.org/index.asp?PAGE=MatchMate">here</a> for CGVA's Match Mate*</font>
		<p />
		<font class='cfont10'><b><asp:label runat="server" ID="messageLabel"></asp:label></b></font>
		</div>

		<form id="Form1" name='INSERT' runat="server">

    <asp:Panel ID="teamPanel" runat="server" >
		<asp:Table ID="teamTable" align='center' cellpadding='3' runat="server">
        
        <asp:tablerow>
        <asp:tablecell colspan="4" bgcolor="#CCCCCC" align="center"><font class="cfont14"><b>My Teams</b></font></asp:tablecell>
        </asp:tablerow>
        
        <asp:tablerow bgcolor="#C0C0C0">
        <asp:tablecell align="center"><font class='cfont10'>Event</font></asp:tablecell>
        <asp:tablecell align="center"><font class='cfont10'>Team Name</font></asp:tablecell>
        <asp:tablecell align="center"><font class='cfont10'>Team Captain</font></asp:tablecell>
        <asp:tablecell align="center"><font class='cfont10'>Captain Email</font></asp:tablecell>
        </asp:tablerow>
        
       </asp:table>
       
       <p />
       <hr color="#000000" />
    </asp:Panel>

    <asp:Panel ID="createTeamPanel" runat="server">
		<table cellspacing='0' align='center' cellpadding='3'>
        
        <tr>
        <td colspan="5" align="center"><font class='cfont10'>Create A New Team For The Current Season (*if you are the team captain*)</font></td>
        </tr>
        
        <tr>
        <td><font class='cfont10'>Event:</font></td>
        <td>
            <asp:DropDownList ID="eventCodeDropDownList" runat="server"></asp:DropDownList>
        </td>
        <td><font class='cfont10'>Team Name:</font></td>
        <td>
            <asp:TextBox ID="teamNameTextBox" runat="server" maxlength="50" size="30"></asp:TextBox>
            </td>
        <td>
            <asp:Button ID="submitButton" runat="server" Text="Create Team" />
            </td>
        </tr>
        
        <tr height="200">
        <td colspan="5">
            <p />

        </td></tr>
        </table>

    </asp:Panel>

		</form>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>


</body>
</html>



