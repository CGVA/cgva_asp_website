<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyEvents.aspx.vb" Inherits="MyEvents" Classname="MyEvents" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">

<head id="Head1" runat="server">
    <title>MyCGVA - My Events</title>
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
	<td height="20" class="cfont16">MyCGVA - My Events</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

		<div align='center'>
		<font class='cfont10'><b>
		Please check the event(s) for which you would like to register, and click 'Continue'.
		</b></font>
		<p />
		 <font class="cfontSuccess10">
		* The registration fee for CGVA leagues can be paid via PayPal, even if you don’t have a PayPal account. *
		</font>
		</div>



        <form name='INSERT' runat="server">

		<table cellspacing='0' align='center' cellpadding='3'>

		<tr>
		<td valign='middle'><font class='cfontError10'><asp:Label ID="messageLabel" runat="server" Text=""></asp:Label></font></td>
		</tr>

		<tr>
		<td>
            <asp:table id="eventTable" runat="server">
		
	    	</asp:table>
        </td>
        </tr>
		</table>

		</form>
		
		<!--JPC 5/2/12
		<div align="center">
		<font class='cfont10'><b>
		Interested in playing in the Spring 2012 Masters Division (Age 40+)?<br />Please <a href="mailto:gabordelon@yahoo.com">email</a> Greg Bordelon with your intent to play or with any questions.
        </b></font>
        </div>
        -->
        
</td>
</tr>

<!-- #include virtual="/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



