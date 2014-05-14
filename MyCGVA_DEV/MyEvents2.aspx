<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyEvents2.aspx.vb" Inherits="MyEvents2" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - My Events - Verify Personal Information</title>

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
	<td height="20" class="cfont16">MyCGVA - My Events - Verify Personal Information</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>


    <tr bgcolor="#FFFFFF">
    <td>

		<div align='center'>
		<font class='cfont10'><b>
        Please verify that your personal information is accurate.<br />
        If anything needs updating, please click <a href="MyProfile.aspx">here</a> before continuing with registration.<br />
        If all information is correct, check the verification checkbox, and click 
        &#39;Continue&#39;. 
        </b></font>
		&nbsp;</div>

		<form id="Form1" name='INSERT' runat="server">
        
        <asp:table id="personInformation" runat="server" align="center">
        
        <asp:tablerow height="50" valign="top" id="messageRow" runat="server">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell colspan="5" align="center"><asp:Label runat="server" id="messageLabel" Text=""></asp:Label></asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
        
        <asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
        <font size='2' face='arial'>First Name<font class="cfontError10">*</font></font><br />
        <asp:TextBox id="FIRST_NAME" size='30' maxlength='25' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
        <font size='2' face='arial'>Last Name<font class="cfontError10">*</font></font><br />
        <asp:TextBox id="LAST_NAME" size='30' maxlength='35' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
        <font size='2' face='arial'>Email<font class="cfontError10">*</font></font><br />
        <asp:TextBox id="EMAIL" size='40' maxlength='50' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>

		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell>
		    <font size='2' face='arial'>Address Line 1</font><br />
            <asp:TextBox id="ADDRESS_LINE1" size='35' maxlength='40' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Address Line 2</font><br />
            <asp:TextBox id="ADDRESS_LINE2" size='35' maxlength='40' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		</asp:tablerow>

		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell>
		    <font size='2' face='arial'>City</font><br />
		    <asp:TextBox id="CITY" size='25' maxlength='30' runat="server" enabled="false"></asp:TextBox>
		</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell>
		    <font size='2' face='arial'>State</font><br />
            <asp:DropDownList id="STATE" runat="server" enabled="false">
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
		</asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Zip Code</font><br />
            <asp:TextBox id="ZIP" size='10' maxlength='10' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		</asp:tablerow>

        <asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Primary Phone (numbers only)</font><br />
            <asp:TextBox id="PRIMARY_PHONE_NUM" size='15' maxlength='10' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>2nd Phone (numbers only)</font><br />
            <asp:TextBox id="PHONE2" size='15' maxlength='10' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>3rd Phone (numbers only)</font><br />
            <asp:TextBox id="PHONE3" size='15' maxlength='10' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>

        <asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Emergency Contact First Name</font><br />
            <asp:TextBox id="EMERGENCY_FIRST_NAME" size='25' maxlength='25' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Emergency Contact Last Name</font><br />
           <asp:TextBox id="EMERGENCY_LAST_NAME" size='25' maxlength='35' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Emergency Phone (numbers only)</font><br />
           <asp:TextBox id="EMERGENCY_PHONE" size='10' maxlength='10' runat="server" enabled="false"></asp:TextBox>
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
        		
        <asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>T-Shirt Size</font><br />
            <asp:DropDownList id="TSHIRT_SIZE" runat="server" enabled="false">
            <asp:ListItem Value="">-Select-</asp:ListItem>
            <asp:ListItem Value="S">Adult-Small</asp:ListItem>
            <asp:ListItem Value="M">Adult-Medium</asp:ListItem>
            <asp:ListItem Value="L">Adult-Large</asp:ListItem>
            <asp:ListItem Value="XL">Adult-XL</asp:ListItem>
            <asp:ListItem Value="XXL">Adult-XXL</asp:ListItem>
            </asp:DropDownList>
         </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Gender</font><br />
            <asp:DropDownList id="GENDER" runat="server" enabled="false">
            <asp:ListItem Value="">-Select-</asp:ListItem>
            <asp:ListItem Value="M">Male</asp:ListItem>
            <asp:ListItem Value="F">Female</asp:ListItem>
            </asp:DropDownList>        
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell>
            <font size='2' face='arial'>Birth Date (mm/dd/yyyy)</font><br />
           <asp:TextBox id="BIRTH_DATE" size='10' maxlength='10' runat="server" enabled="false"></asp:TextBox>
            <br />
            <font class='cfont10'> </font><font class='cfontError10'><asp:CompareValidator id="cvStartDate" runat="server" ControlToValidate="BIRTH_DATE"
            Operator="DataTypeCheck" Type="Date" text="" ErrorMessage="Birth date is invalid." EnableClientScript="True"></asp:CompareValidator></font>
        </asp:tablecell>        
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
        
    	<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell colspan='5'>
		    <font size='2' face='arial'>How did you hear about CGVA?<font class="cfontError10">*</font></font><br />
            <asp:DropDownList id="FIRST_CONTACT_ID" runat="server" enabled="false">
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
                 </asp:DropDownList>
        </asp:tablecell>
         <asp:tablecell width="10">&nbsp;</asp:tablecell>
       </asp:tablerow>

		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell colspan='5'><font size='2' face='arial'>Would you like to receive email about CGVA events?</font><br />
            <!--reverse these because of how the question is asked-->
            <asp:DropDownList id="SUPPRESS_EMAIL_IND" runat="server" 
            enabled="false">
            <asp:ListItem Value="Y">No</asp:ListItem>
            <asp:ListItem Value="N">Yes</asp:ListItem>
            </asp:DropDownList>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>

		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell colspan='5'>
		    <font size='2' face='arial'>Would you like to receive postal mail about CGVA events?</font><br />
            <!--reverse these because of how the question is asked-->
            <asp:DropDownList id="SUPPRESS_SNAIL_MAIL_IND" runat="server" enabled="false">
                <asp:ListItem Value="Y">No</asp:ListItem>
                <asp:ListItem Value="N">Yes</asp:ListItem>
            </asp:DropDownList> <font size='1' face='arial'>(CGVA will rarely send mail to your home address)</font>
		</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		</asp:tablerow>

		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell colspan="5">
		    <font size='2' face='arial'>Would you like to suppress your last name from the CGVA website?</font><br />
		    <asp:DropDownList id="SUPPRESS_LAST_NAME_IND" runat="server" enabled="false">
                <asp:ListItem Value="N">No</asp:ListItem>
                <asp:ListItem Value="Y">Yes</asp:ListItem>
            </asp:DropDownList> <font size='1' face='arial'>(‘Yes’ = only last initial will display)</font>
        </asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
    
		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell colspan="5">
            <font size='2' face='arial'>NAGVA Rating (if you have one)</font><br />
		    <asp:TextBox id="NAGVA_RATING" size='10' maxlength='50' runat="server" enabled="false"></asp:TextBox>
		</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
		
		<asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        <asp:tablecell colspan="5">
            <font size='2' face='arial'>CGVA Rating</font><br />
		    <font size='2' face='arial'><asp:label id="ratingLabel" runat="server"></asp:label></font>
		</asp:tablecell>
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
        </asp:tablerow>
        
        <asp:tablerow height="50" valign="top">
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		<asp:tablecell colspan='5' align='right' valign='middle'>
		    <font class='cfont10'>
            <asp:CheckBox id="verifyInfo" runat="server" text="All information is correct" forecolor="Black" />
            &nbsp;
           <asp:Button id="submitButton" runat="server" text="Continue"  /></font>
        </asp:tablecell>                
        <asp:tablecell width="10">&nbsp;</asp:tablecell>
		</asp:tablerow>

        </asp:table>

		</form>

</td>
</tr>

<!-- #include virtual="/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



