<%@ Page Language="VB" AutoEventWireup="false" CodeFile="NewRegistration.aspx.vb" Inherits="NewRegistration" enableEventValidation="false" %>
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
	<td height="20" class="cfont16">MyCGVA - New Member Registration</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>



<tr bgcolor="#FFFFFF">
<td>

		<div align='center'>
		<font class='cfont10'><b>
		Please enter the registration information below. Required fields are marked with an <font class="cfontError10">*</font>.
		<br />After submitting your registration, you will receive an email with your 
            user name and password for MyCGVA.</b></font>
		</div>

		<br />

		        <form id="INSERT" runat="server">

		<table cellspacing='0' align='center' cellpadding='3'>

        <tr>
        <td colspan='2' align='center'>
        <font class="cfontError10">
        <asp:Label ID="messageLabel" runat="server" Text=""></asp:Label>
        </font>
        </td>
        </tr>
		
				<tr>
		<td align='right' valign='middle' class="style1"><font size='2' face='arial'>First Name<font class="cfontError10">*</font></font></td>
		<td class="style1"><asp:TextBox ID="FIRST_NAME" size='25' maxlength='25' runat="server"></asp:TextBox>
            </td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Last Name<font class="cfontError10">*</font></font></td>
		<td><asp:TextBox ID="LAST_NAME" size='25' maxlength='35' runat="server"></asp:TextBox>
                        </td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Email<font class="cfontError10">*</font></font></td>
		<td><asp:TextBox ID="EMAIL" size='40' maxlength='50' runat="server"></asp:TextBox>
                        </td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>How did you hear about CGVA?<font class="cfontError10">*</font></font></td>
		<td>
        <asp:DropDownList ID="FIRST_CONTACT_ID" runat="server">
			 <asp:ListItem value='0'>-select-</asp:ListItem>
			 <asp:ListItem value='6'>CGVA Website</asp:ListItem>
			 <asp:ListItem value='2'>Email</asp:ListItem>
			 <asp:ListItem value='8'>Friends</asp:ListItem>
			 <asp:ListItem value='523'>League/Tournament Sign Up</asp:ListItem>
			 <asp:ListItem value='10'>Other</asp:ListItem>
			 <asp:ListItem value='9'>Other Publication</asp:ListItem>
			 <asp:ListItem value='3'>Out Front Ad</asp:ListItem>
			 <asp:ListItem value='4'>Poster</asp:ListItem>
			 <asp:ListItem value='7'>Pride Fest Booth</asp:ListItem>
			 <asp:ListItem value='1'>Unknown</asp:ListItem>
			 <asp:ListItem value='5'>US Mail</asp:ListItem>
</asp:DropDownList>		</td>
		</tr>
		
		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>T-Shirt Size</font></td>
		<td>
        <asp:DropDownList ID="TSHIRT" runat="server">
                <asp:ListItem Value="">-Select-</asp:ListItem>
                <asp:ListItem Value="S">Adult-Small</asp:ListItem>
                <asp:ListItem Value="M">Adult-Medium</asp:ListItem>
                <asp:ListItem Value="L">Adult-Large</asp:ListItem>
                <asp:ListItem Value="XL">Adult-XL</asp:ListItem>
                <asp:ListItem Value="XXL">Adult-XXL</asp:ListItem>
            </asp:DropDownList>
            </td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Gender</font></td>
		<td>
<asp:DropDownList ID="GENDER" runat="server">
                <asp:ListItem Value="">-Select-</asp:ListItem>
                <asp:ListItem Value="M">Male</asp:ListItem>
                <asp:ListItem Value="F">Female</asp:ListItem>
            </asp:DropDownList>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Primary Phone 
            Number</font></td>
		<td><asp:TextBox ID="PRIMARY_PHONE_NUM" size='10' maxlength='10' runat="server"></asp:TextBox>
                        <font size='2' face='arial'>numbers only)</font></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Secondary Phone Number</font></td>
		<td><asp:TextBox ID="PHONE2" size='10' maxlength='10' runat="server"></asp:TextBox>
                        <font size='2' face='arial'>(numbers only)</font></td>
		</tr>

        <tr>
        <td bgcolor="#ffffff">
        &nbsp;</td>
        <td align="right" bgcolor="#ffffff" valign="top">
        &nbsp;</td>
        </tr>

        <tr>
        <td  align="right" bgcolor="#ffffff">
        <font class="cfont12"><b>Create Password</b></font><font class="cfontError10">*</font></td>
        <td align="right" bgcolor="#ffffff" valign="top">
        &nbsp;</td>
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
            <asp:Button ID="submitButton" runat="server" Text="Submit New Member Registration Request" />
            </font></td>
		</tr>


		
		</table>

		        </form>

</td>
</tr>

<!-- #include virtual="~/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



