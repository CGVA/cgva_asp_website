<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyEvents5.aspx.vb" Inherits="MyEvents5" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - My Events - Payment Information</title>

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
	<td height="20" class="cfont16">MyCGVA - My Events - Payment Information</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

		<tr>
		<td colspan='2' align='center'><font class="cfont14"><b><u>My Events - Payment Information</u></b></font></td>
		</tr>

        <tr bgcolor="#FFFFFF">
        <td>

		<div align='center'>
		<font class='cfont10'><b>
		Click on the PayPal button to make your payment.</b></font>
		<p />
        <!-- LEAGUES -->
<%--        <font class='cfontError14'>***IMPORTANT!***</font>
        <br /><br />
        <font class='cfont10'><b>
        On the final Paypal screen, 
        <br />
        please make sure to click on the 'Return To CGVA' button for your registration to be completed.</b></font>
        <br />
        <br />
--%>
        
        <!-- LEAGUES -->
        
        <!-- VOLLEYPALOOZA ONLY -->
        
		<font class='cfont10'><b>
		Choose how many players you are paying for, and then click the PayPal button to make 
		your payment.
		<br />
        <br />
        <font class='cfontError14'>***IMPORTANT!***</font>
        <br /><br />
            If you are paying for additional players, please include the player names in the 
            &#39;Add Special Instructions To Merchant&#39; field on the payment screen.</b></font><p />
        <!-- VOLLEYPALOOZA ONLY -->
        
        <form runat="server" id="form1">
        
        <table cellspacing='0' align='center' cellpadding='3'>
        <tr>
        <td valign="middle">
        <font class='cfont10'>Event(s) chosen:
        &nbsp;
    <!--    use this if there is to be more than one value in the DDL -->
        <asp:DropDownList ID="leagueChoiceDropDownList" runat="server" autopostback="true"></asp:DropDownList>

    <!--    use this if there is to be more than one value in the DDL -->
<%--        
        <asp:DropDownList ID="leagueChoiceDropDownList" runat="server" disabled="true"></asp:DropDownList>
--%>            
        </font>   
        </td></tr>
        
                <tr>
        <td align="right">
            <asp:ImageButton ID="paypalImageButton" runat="server" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" runat="server" />
            <input type="hidden" name="hosted_button_id" value="2055">
            <input type="hidden" name="currency_code" value="USD">
            <input type="hidden" name="cmd" value="_xclick">
            <asp:HiddenField ID="amount" value="" runat="server" />
            <asp:HiddenField ID="item_name" value="2014 Fall League" runat="server" />
            <asp:HiddenField ID="return" value="http://cgva.org/MyCGVA/MyEvents6.aspx" runat="server" />


        </td></tr>
        
        <tr><td height="200">&nbsp;</td></tr>
        
        </table>
        
        <asp:HiddenField ID="verifyWaiverInfo" runat="server" />
<asp:Label runat="server" ID="lblText" />
    </form>


</td>
</tr>

<!-- #include virtual="/incs/fragSiteContact.asp" -->

</table>

</body>
</html>




