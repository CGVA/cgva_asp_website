﻿<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Registration.aspx.vb" Inherits="Registration" enableEventValidation="false" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html40/strict.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="author" content="John Cross">
<meta name="description" content="Tournament information for Last Dig in Denver hosted by CGVA.">
<meta name="keywords" content="CGVA, Volleyball, Last Dig, Denver, Gay, NAGVA, Tournament">
<title>Home - Last Dig in Denver</title>
<link rel="stylesheet" type="text/css" media="all" href="main.css">
<link rel="stylesheet" type="text/css" media="all" href="colorschemes/colorscheme1/colorscheme.css">
<link rel="stylesheet" type="text/css" media="all" href="style.css">
<script type="text/javascript" src="live_tinc.js"></script>
</head>
<body id="main_body">
<div id="page_all">
<div id="container">
<div id="main_nav_container">
<ul id="main_nav_list">
<li><a class="main_nav_item" href="index.html">Home</a></li>
<li><a class="main_nav_active_item" href="registration.aspx">Registration</a></li>
<li><a class="main_nav_item" href="Schedule.html">Schedule</a></li>
<li><a class="main_nav_item" href="facility.html">Facility</a></li>
<li><a class="main_nav_item" href="Hotel.html">Hotel</a></li>
<li><a class="main_nav_item" href="sponsors.html">Sponsors</a></li>
<li><a class="main_nav_item" href="aboutus.html">About Us</a></li>
</ul>
</div>

<div id="main_container">
<div id="key_visual">&nbsp;</div>
<div id="logo"></div>
<div id="sub_container1">
    <div id="sub_nav_container"></div>
</div>
<div id="sub_container2">
    <div id="slogan"></div>
<div class="content2" id="content_container">

<span style="font-size:14px;font-weight:bold"><asp:Label ID="successLabel" runat="server" Text=""></asp:Label></span>

 <h3>Registration:</h3>
<span style="font-size:14px;">
You will need to register your team on the <a target="_blank" title="NAGVA website" href="http://www.nagva.org">NAGVA</a>
 website and then send in your team fees.  Please note that registration is based on when your payment is received, not when 
 you submit your roster online.
</span>
<br />

<h3>Fees:</h3>
<p>
<span style="font-size:14px;">
    Team fee are based on date of registration for up to 7 players, and additional players may be registered 
    for $65 per player up to the normal registration cut off date.  Late registration, after 3/25, for 
    additional players is $70 per player. Registration fees and dates: $420 by 3/1/11, $455 by 3/25/11. 
    Late Registration of $505 by 4/1/11.
</span></p>

<p>
<span style="font-size:14px;">All fees are in U.S. currency and must be paid by 
    either a Cashier&#39;s Check or Money Order payable to &quot;<span style="font-weight:bold;">CGVA</span>&quot;.&nbsp; 
    We are sorry, but Personal Checks will not be accepted.</span>
</p>

<h3>Pay with Paypal</h3>

<asp:Panel ID="closedPanel" runat="server">
<p>
<span style="font-size:14px;">Paypal registration is closed at this time.</span>
</p>
</asp:Panel>

<asp:Panel ID="openPanel" runat="server">

<span style="font-size:14px;">Payments made through PayPal will be charged a $15 fee 
    for the team and $3 for extra players in addition to the normal registration to 
    cover the cost of the transaction.
</span>
    
<!-- 
<p>
JPC 1/29/11
<span style="font-size:14px;">Payments made through PayPal will be charged $15 fee for the team and $3 for extra players in addition to the normal registration to cover the cost of the transaction.
Paypal is available for a small $15 fee for 6 or 7 
    member teams and $3 per additional player.</span>
</p>
-->

<form name="form1" id="form1" runat="server">
<table border="0" cellspacing='0' cellpadding='3'>
<tr>
<td>
    <h4>New Team Registration</h4>
    <span style="font-size:14px;">
    Enter your team name and select your division* (e.g., Team XYZ - BB)
    <br />
    If your team has more than 7 players, select the number of additional players
    <br />
    Proceed to PayPal checkout
    </span>
    <p />
</td>
</tr>
    
    <tr>
    <td><b>
        $<asp:Label ID="teamFeeLabel" runat="server" Text="XXX"></asp:Label> per team of up to 7 players, if paid by 
        <asp:Label ID="registrationDateLabel" runat="server" Text="XXX"></asp:Label></b><br /></td>
    </tr>

    <tr>
    <td>
        <span style="font-size:14px;font-weight:bold;color:#FF0000">
        <asp:Label ID="errorLabelTeam" runat="server" Text=""></asp:Label> 
        </span>
        
        <table>
        <tr>
        <td>Team Name <font color="red">*</font></td>
        <td>
            <asp:TextBox ID="teamName" maxlength="50" size="30" runat="server"></asp:TextBox></td>
        <td>Division <font color="red">*</font></td>
        <td>
            <asp:DropDownList id="divisionDDL" runat="server">
            <asp:ListItem value=''>-select-</asp:ListItem>
            <asp:ListItem value='B'>B</asp:ListItem>
            <asp:ListItem value='BB'>BB</asp:ListItem>
            <asp:ListItem value='Modified A'>Modified A</asp:ListItem>
            </asp:DropDownList>
        </td>
        </tr>
        </table>
        <br />
    </td>
    </tr>
    
    <tr>
    <td><b>$<asp:Label ID="additionalPayerFeeLabel" runat="server" Text="XXX"></asp:Label> per additional player for a team of 8 or 
        more</b><br /></td>
    </tr>

     <tr>
    <td>

        <table>
        <tr>
        <td>Additional Players</td>
        <td>
            <asp:DropDownList id="additionalPlayersDDL" runat="server">
			<asp:ListItem value='0'>-select-</asp:ListItem>
            </asp:DropDownList>
        </td>
        </tr>
        </table>
        <br />
    </td>
    </tr>

<tr>
<td>
<asp:HiddenField id="registrationType" value="TEAM" runat="server" />
<asp:HiddenField id="teamFee" runat="server" />
<asp:ImageButton  ID="paypalImageButton" runat="server" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" runat="server" />
</td>
</tr>
    
<tr><td><font color="red">* Required.</font></td></tr>

    </table>

<table border="0" cellspacing='0' cellpadding='3'>
<tr>
<td>
<h4>Existing Team Update</h4>
<span style="font-size:14px;">
Teams will not be available in the drop down list to add extra players until the original PayPal payment is processed.  
If you need to add additional players and your team is not visible, please contact the Registrar.
<br /><br />
To add players:
<br />
Select your team name/division
<br />
    If your team has more than 7 players, select the number of additional players
<br />
    Proceed to PayPal checkout
</span>
<p />
</td>
</tr>
 <tr>
<td>
 
         <span style="font-size:14px;font-weight:bold;color:#FF0000">
        <asp:Label ID="errorLabelAdditionalPlayers" runat="server" Text=""></asp:Label> 
        </span>
        
    <table>
    <tr>
            <td>Team Name <font color="red">*</font></td>
    <td>
             <asp:DropDownList id="teamNameDDL" runat="server">
			<asp:ListItem value=''>-select-</asp:ListItem>
            </asp:DropDownList>
    </td>
    <td>Additional Players <font color="red">*</font></td>
    <td>
            <asp:DropDownList id="additionalPlayersDDL2" runat="server">
			<asp:ListItem value='0'>-select-</asp:ListItem>
            </asp:DropDownList>
    </td>
    </tr>
    </table>
    <br />
</td>
</tr>

<tr>
<td>
<asp:HiddenField ID="registrationType2" value="PLAYER" runat="server" />
<asp:ImageButton ID="paypalImageButton2" runat="server" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" runat="server" />
</td>
</tr>
    
<tr><td><font color="red">* Required.</font></td></tr>
</table>
</form>    
</asp:Panel>



 
<h3>Payment by mail</h3>
<p><span style="font-size:14px;">
<span style="font-weight:bold;"><span style="font-size:14px ! important;">Send 
    Payments to:</span></span>
<br />
<span style="font-size:14px;">
<span style="font-size:14px ! important;">CGVA - Last Dig in Denver</span>
<br />
    Po Box 18576
<br />
    Denver, CO 80218-0576
</span> 
<br />
</span>
</p>

<h3>Registration Deadline:</h3>
<span style="font-size:14px;">
<span style="font-size:14px ! important;">
Normal registration closes on March 25th, 2011.  Late registration closes on April 1st, 2011.  
Entries must be postmarked by the given dates to be accepted.  If you are submitting a late registration, 
you MUST email the tournament director to let them know to expect your entry. 

</span></span>
<br />

<h3>Refund Policy:</h3>
<span style="font-size:14px;">
CGVA will refund 100% of your registration fee if requested on or before March 19, 2011
<br />
CGVA will refund 50% of your registration fee if requested between  March 20, 2011 and March 25, 2011
<br />
No refunds will be given if requested on or after March 26, 2011
<br />
</span>
</div>
</div>

<div id="footer_text">Modified 1/17/11</div>
</div>
</div>
</div>
<p><!-- wfxbuild / 1.0 / layout6-15-l1 / 2010-01-07 08:02:55 CET--></p>
       
     
</body></html>