<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyProfileTest.aspx.vb" Inherits="MyProfileTest" enableEventValidation="false" %>
<% System.Web.HttpContext.Current.Response.AddHeader( "Cache-Control","no-cache")
System.Web.HttpContext.Current.Response.Expires = 0
    System.Web.HttpContext.Current.Response.Cache.SetNoStore()
System.Web.HttpContext.Current.Response.AddHeader( "Pragma", "no-cache")%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - My Profile</title>

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
    Response.WriteFile("/incs/fragMyCGVANavigation.asp")
%>
</td>
<td>&nbsp;</td>
<td valign='top' width='100%'>


	<table width='100%' bgcolor="#FFFFFF" border="0">

	<tr bgcolor="#FF3300">
	<td height="20" class="cfont16">MyCGVA - My Profile</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

     <tr>
     <td valign='top'>

		<div align='center'>
		<font class='cfont10'><b>
		Please add/update and of your personnel information below. Required fields are marked with an <font class="cfontError10">*</font>.
		</b></font>
		</div>

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
		<td><asp:TextBox ID="FIRST_NAME" size='25' maxlength='25' runat="server"></asp:TextBox>
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
        <asp:DropDownList ID="TSHIRT_SIZE" runat="server">
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
		<td align='right' valign='middle'><font size='2' face='arial'>Birth Date</font></td>
		<td><asp:TextBox ID="BIRTH_DATE" size='10' maxlength='10' runat="server"></asp:TextBox>
		<font class='cfont10'> (mm/dd/yyyy)</font><font class='cfontError10'><asp:CompareValidator id="cvStartDate" runat="server" ControlToValidate="BIRTH_DATE"
Operator="DataTypeCheck" Type="Date" Text="" ErrorMessage="Birth date is invalid." SetFocusOnError="true" EnableClientScript="False"></asp:CompareValidator></font>
</td>
		</tr>


		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Primary Phone 
            Number</font></td>
		<td><asp:TextBox ID="PRIMARY_PHONE_NUM" size='10' maxlength='10' runat="server"></asp:TextBox>
                        <font size='2' face='arial'> (numbers only)</font></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Secondary Phone Number</font></td>
		<td><asp:TextBox ID="PHONE2" size='10' maxlength='10' runat="server"></asp:TextBox>
                        <font size='2' face='arial'> (numbers only)</font></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>3rd Phone Number</font></td>
		<td><asp:TextBox ID="PHONE3" size='10' maxlength='10' runat="server"></asp:TextBox>
                        <font size='2' face='arial'> (numbers only)</font></td>
		</tr>
		
		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Address Line 1</font></td>
		<td><asp:TextBox ID="ADDRESS_LINE1" size='35' maxlength='40' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Address Line 2</font></td>
		<td><asp:TextBox ID="ADDRESS_LINE2" size='35' maxlength='40' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>City</font></td>
		<td><asp:TextBox ID="CITY" size='25' maxlength='30' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>State</font></td>
		<td>
            <asp:DropDownList ID="STATE" runat="server">
			<asp:ListItem value=''>-select-</asp:ListItem>
			<asp:ListItem value='AK'>Alaska</asp:ListItem>
			<asp:ListItem value='AL'>Alabama</asp:ListItem>
			<asp:ListItem value='AZ'>Arizona</asp:ListItem>
			<asp:ListItem value='AR'>Arkansas</asp:ListItem>
			<asp:ListItem value='CA'>California</asp:ListItem>
			<asp:ListItem value='CO' selected>Colorado</asp:ListItem>
			<asp:ListItem value='CT'>Connecticut</asp:ListItem>
			<asp:ListItem value='DC'>District of Columbia</asp:ListItem>
			<asp:ListItem value='DE'>Deleware</asp:ListItem>
			<asp:ListItem value='FL'>Florida</asp:ListItem>
			<asp:ListItem value='GA'>Georgia</asp:ListItem>
			<asp:ListItem value='HI'>Hawaii</asp:ListItem>
			<asp:ListItem value='ID'>Idaho</asp:ListItem>
			<asp:ListItem value='IL'>Illinois</asp:ListItem>
			<asp:ListItem value='IN'>Indiana</asp:ListItem>
			<asp:ListItem value='IA'>Iowa</asp:ListItem>
			<asp:ListItem value='KS'>Kansas</asp:ListItem>
			<asp:ListItem value='KY'>Kentucky</asp:ListItem>
			<asp:ListItem value='LA'>Lousiana</asp:ListItem>
			<asp:ListItem value='ME'>Maine</asp:ListItem>
			<asp:ListItem value='MD'>Maryland</asp:ListItem>
			<asp:ListItem value='MA'>Massachussetts</asp:ListItem>
			<asp:ListItem value='MI'>Michigan</asp:ListItem>
			<asp:ListItem value='MN'>Minnesota</asp:ListItem>
			<asp:ListItem value='MS'>Mississippi</asp:ListItem>
			<asp:ListItem value='MO'>Missouri</asp:ListItem>
			<asp:ListItem value='MT'>Montana</asp:ListItem>
			<asp:ListItem value='NE'>Nebraska</asp:ListItem>
			<asp:ListItem value='NV'>Nevada</asp:ListItem>
			<asp:ListItem value='NH'>New Hampshire</asp:ListItem>
			<asp:ListItem value='NJ'>New Jersey</asp:ListItem>
			<asp:ListItem value='NM'>New Mexico</asp:ListItem>
			<asp:ListItem value='NY'>New York</asp:ListItem>
			<asp:ListItem value='NC'>North Carolina</asp:ListItem>
			<asp:ListItem value='ND'>North Dakota</asp:ListItem>
			<asp:ListItem value='OH'>Ohio</asp:ListItem>
			<asp:ListItem value='OK'>Oklahoma</asp:ListItem>
			<asp:ListItem value='OR'>Oregon</asp:ListItem>
			<asp:ListItem value='PA'>Pennsylvania</asp:ListItem>
			<asp:ListItem value='RI'>Rhode Island</asp:ListItem>
			<asp:ListItem value='SC'>South Carloina</asp:ListItem>
			<asp:ListItem value='SD'>South Dakota</asp:ListItem>
			<asp:ListItem value='TN'>Tennessee</asp:ListItem>
			<asp:ListItem value='TX'>Texas</asp:ListItem>
			<asp:ListItem value='UT'>Utah</asp:ListItem>
			<asp:ListItem value='VE'>Vermont</asp:ListItem>
			<asp:ListItem value='VA'>Virginia</asp:ListItem>
			<asp:ListItem value='WA'>Washington</asp:ListItem>
			<asp:ListItem value='WV'>West Virginia</asp:ListItem>
			<asp:ListItem value='WI'>Wisconsin</asp:ListItem>
			<asp:ListItem value='WY'>Wyoming</asp:ListItem>
			</asp:DropDownList>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Zip Code</font></td>
		<td><asp:TextBox ID="ZIP" size='10' maxlength='10' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Emergency Contact First Name</font></td>
		<td><asp:TextBox ID="EMERGENCY_FIRST_NAME" size='25' maxlength='25' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Emergency Contact Last Name</font></td>
		<td><asp:TextBox ID="EMERGENCY_LAST_NAME" size='25' maxlength='35' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Emergency Phone 
            Number</font></td>
		<td><asp:TextBox ID="EMERGENCY_PHONE" size='10' maxlength='10' runat="server"></asp:TextBox><font size='2' face='arial'> (numbers only)</font></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Would you like to receive email about CGVA events?</font></td>
		<td>
            <!--reverse these because of how the question is asked-->
            </reverse><asp:DropDownList ID="SUPPRESS_EMAIL_IND" runat="server">
                <asp:ListItem Value="Y">No</asp:ListItem>
                <asp:ListItem Value="N">Yes</asp:ListItem>
            </asp:DropDownList>
            
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Would you like to receive postal mail about CGVA events?</font></td>
		<td>
            <!--reverse these because of how the question is asked-->
            <asp:DropDownList ID="SUPPRESS_SNAIL_MAIL_IND" runat="server">
                <asp:ListItem Value="Y">No</asp:ListItem>
                <asp:ListItem Value="N">Yes</asp:ListItem>
            </asp:DropDownList> <font size='1' face='arial'>(CGVA will rarely send mail to your home address)</font>
		</td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Would you like to suppress your last name from the CGVA website?</font></td>
		<td>
            <asp:DropDownList ID="SUPPRESS_LAST_NAME_IND" runat="server">
                <asp:ListItem Value="N">No</asp:ListItem>
                <asp:ListItem Value="Y">Yes</asp:ListItem>
            </asp:DropDownList> <font size='1' face='arial'>(‘Yes’ = only last initial will display)</font></td>
		</tr>
    
		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>NAGVA Rating (if you have one)</font></td>
		<td><asp:TextBox ID="NAGVA_RATING" size='10' maxlength='50' runat="server"></asp:TextBox></td>
		</tr>

		<tr>
		<td align='right' valign='middle'><font size='2' face='arial'>Draft League Picture</font></td>
		<td><font size='2' face='arial'>
            <asp:FileUpload ID="PHOTO_NAME" size='40' runat="server" />
            </font></td>
		</tr>
		
        <tr>
        <td align='right' valign='top'><font class="cfont10">
        <asp:label text="Picture on file" runat="server" ID="lblImage"></asp:label>
        </font></td>
        <td><font class="cfont10">
        <asp:label text="" runat="server" ID="lblImageFound"></asp:label>
        <asp:image ID="imgFile" runat="server"></asp:image>
        </font></td>
        </tr>

		<tr>
		<td colspan='2' align='right'><font size="3" face='arial'><asp:Button ID="submitButton" runat="server" Text="Update My Profile" /></font></td>
		</tr>

		</table>

		</form>

</td>
</tr>

<%		   
    Response.WriteFile("/incs/fragSiteContact.asp")
%>

</table>

</body>
</html>



