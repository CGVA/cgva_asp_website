﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="IPNTEST.aspx.vb" Inherits="IPNTEST" enableEventValidation="false" %>
<%  Response.Expires = -1
    Response.Clear()
    Response.CacheControl = "no-cache"%>
    <html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - My Events - Payment Success</title>

<%		   
    Response.WriteFile("/incs/fragHeader.asp")
%>

</head>
<%		   
    Response.WriteFile("/incs/Header.asp")
%>
<form id="form1" runat="server">
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
	<td height="20" class="cfont16">MyCGVA - My Events - Registration Complete</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

	<tr>
	<td colspan='2' align='center'><font class="cfont14"><b><u>My Events - Registration Complete</u></b></font></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td height="500" valign='top'>

	<div align='center'>
	    <br /><br />
	    <font class='cfontSuccess10'>
        <asp:Label ID="messageLabel" runat="server" Text=""></asp:Label>
	    </font>
   	<br />
    </div>
    </td></tr>
    
   </table>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</form>

</html>