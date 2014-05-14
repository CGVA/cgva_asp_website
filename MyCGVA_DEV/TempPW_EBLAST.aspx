<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TempPW_EBLAST.aspx.vb" Inherits="TempPW_EBLAST" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
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
	<td height="20" class="cfont16">MyCGVA - Temp Email E-blast</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>



<tr bgcolor="#FFFFFF">
<td>


		<form id="INSERT" runat="server">

		<table cellspacing='0' align='center' cellpadding='3'>

        <tr>
        <td colspan='2' align='center'>
        <font class="cfontError10">
        <asp:Label ID="messageLabel" runat="server" Text=""></asp:Label>
        </font>
        </td>
        </tr>
		
		</table>
        </form>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</html>




