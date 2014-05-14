<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyEvents4.aspx.vb" Inherits="MyEvents4" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <!--League Information/Skills Development/Player Rating-->
    <title>MyCGVA - My Events - League Information/Player Rating</title>

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
	<!--<td height="20" class="cfont16">MyCGVA - My Events - League Information/Skills Development/Player Rating</td>-->
	<td height="20" class="cfont16">MyCGVA - My Events - League Information/Player Rating</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>


	<tr>
	<td colspan='2' align='center'><font class="cfont14"><b><u>League Information</u></b></font></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

	<div align='center'>
	<font class='cfont10'><b>Please answer the following questions about the league(s) for which you are registering:</b></font>
	<br />
    </div>

	<form name='INSERT' runat="server">

    <asp:Panel ID="Panel1" runat="server" >
		
		<asp:Panel ID="LeaguePanel" runat="server" >
		
		<table cellspacing='0' align='center' cellpadding='3'>
		<tr>
		<td colspan='2' align='center'><font class="cfont10"><b>2011 - Summer League</b></font></td>
		</tr>
		
		<tr>
		<td align="right"><font class='cfont10'>Are you the captain of a team?</font></td>
		<td>
            <asp:DropDownList ID="CAPTAIN_IND_League" runat="server">
                <asp:ListItem Value="N">No</asp:ListItem>
                <asp:ListItem Value="Y">Yes</asp:ListItem>
            </asp:DropDownList>
		</td>
		</tr>
			
		<tr>
		<td colspan='2' align='center'>
            &nbsp;
            </td>
		</tr>
        </table>
        
        </asp:Panel>		
		
   
    </asp:Panel>

    <asp:Panel ID="skillsDevelopmentPanel" runat="server" >
	<table cellspacing='0' align='center' cellpadding='3'>

    <tr>
	<td colspan='2' align='center'>&nbsp;</td>
	</tr>

	<tr>
	<td colspan='2' align='center'><font class="cfont14"><b>Skills Development/Player Rating</b></font></td>
	</tr>
	
	<tr>
	<td align="right"><font class='cfont10'>Would you participate in CGVA Skill Development Sessions?</font></td>
	<td>
        <asp:DropDownList ID="Q1" runat="server">
            <asp:ListItem Value="">-select-</asp:ListItem>
            <asp:ListItem Value="N">No</asp:ListItem>
            <asp:ListItem Value="Y">Yes</asp:ListItem>
        </asp:DropDownList>
	</td>
	</tr>
			
	<tr>
	<td align="right"><font class='cfont10'>If yes, would you prefer individual skills or a team clinic?</font></td>
	<td>
        <asp:DropDownList ID="Q2" runat="server">
            <asp:ListItem Value="">-select-</asp:ListItem>
            <asp:ListItem Value="I">Individual</asp:ListItem>
            <asp:ListItem Value="T">Team</asp:ListItem>
        </asp:DropDownList>
	</td>
	</tr>

	<tr>
	<td align="right"><font class='cfont10'>What skill would you most like to learn/practice?</font></td>
	<td>
	    <asp:TextBox ID="Q3" runat="server" maxlength="50" Width="262px"></asp:TextBox>
	</td>
	</tr>
	
	<tr>
	<td align="right"><font class='cfont10'>What level of player are you?</font></td>
	<td>
        <asp:DropDownList ID="Q4" runat="server">
            <asp:ListItem Value="">-select-</asp:ListItem>
            <asp:ListItem Value="A">Advanced (Rating BB or above)</asp:ListItem>
            <asp:ListItem Value="I">Intermediate (Rating B)</asp:ListItem>
            <asp:ListItem Value="B">Beginner (Unrated/New Player)</asp:ListItem>
        </asp:DropDownList>
	</td>
	</tr>

    <tr>
	<td colspan='2' align='center'>&nbsp;</td>
	</tr>
    
    </table>
    </asp:Panel>

	<table cellspacing='0' align='center' cellpadding='3'>

    <tr>
	<td colspan='2' align='center'>&nbsp;</td>
	</tr>
	
	<tr>
	<td colspan='2' align='center'>
	<font class="cfont14"><b>Player Rating</b></font>
	<!--placeholder for Q1-Q4-->
	<input type="hidden" id="Q1" value="" />
	<input type="hidden" id="Q2" value="" />
	<input type="hidden" id="Q3" value="" />
	<input type="hidden" id="Q4" value="" />
	</td>
	</tr>

	<tr>
	<td valign='middle' align="right"><font class='cfont10'><asp:Label ID="ratingLabel" runat="server" Text=""></asp:Label></font></td>
	<td valign='middle'>
        <asp:DropDownList ID="Q5" runat="server">
            <asp:ListItem Value="N">No</asp:ListItem>
            <asp:ListItem Value="Y">Yes</asp:ListItem>
        </asp:DropDownList>
	</td>
	</tr>
	
	<tr>
	<td colspan='2' align='center'>&nbsp;</td>
	</tr>

		<tr>
		<td colspan="2" align="right"><font class='cfont10'>
            <asp:Button ID="submitButton" runat="server" Text="Continue To PayPal" />
            </font></td>
	</tr>

	</table>
        <asp:HiddenField id="League" runat="server" />
        <asp:HiddenField id="SB" runat="server" />
        <asp:HiddenField id="verifyInfo" runat="server" />
        <asp:HiddenField id="verifyWaiverInfo" runat="server" />
	</form>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



