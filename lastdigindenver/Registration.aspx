<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Registration.aspx.vb" Inherits="Registration" enableEventValidation="false" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html40/strict.dtd">
<html><head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Expires" CONTENT="-1">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
   <meta name="author" content="John Crossin"><meta name="description" content="Tournament information for Last Dig in Denver hosted by CGVA.">
   <meta name="keywords" content="CGVA, Volleyball, Last Dig, Denver, Gay, NAGVA, Tournament">
   <title>Home - Last Dig in Denver</title>
   <link rel="stylesheet" type="text/css" media="all" href="main.css">
   <link rel="stylesheet" type="text/css" media="all" href="colorschemes/colorscheme1/colorscheme.css">
   <link rel="stylesheet" type="text/css" media="all" href="style.css">
   <script type="text/javascript" src="live_tinc.js"></script></head>
   <body id="main_body"><div id="page_all"><div id="container">
   <div id="main_nav_container">
   <ul id="main_nav_list"><li><a class="main_nav_item" href="index.html">Home</a></li>
   <li><a class="main_nav_active_item" href="Registration.aspx">Registration</a></li>

<li><a class="main_nav_item" href="schedule.html">Schedule</a></li>
    <li><a class="main_nav_item" href="poolplay.html">Pool Play/Bracket</a></li>
   <li><a class="main_nav_item" href="facility2.html">Facility/Hotel</a></li>
   <li><a class="main_nav_item" href="sponsors2.html">Sponsors</a></li>
   <li><a class="main_nav_item" href="aboutus.html">About Us</a></li>

   </ul>
   </div>
   <div id="main_container"><div id="key_visual">&nbsp;
    </div>
    <div id="logo"></div>
    <div id="sub_container1">
    <div id="sub_nav_container"></div>
    </div>
    <div id="sub_container2">
    <div id="slogan"></div>
<div class="content2" id="content_container">

<span style="font-size:14px;font-weight:bold"><asp:Label ID="successLabel" runat="server" Text=""></asp:Label></span>

<br />
<h3>Registration</h3>
<!--
<br />
<span style="font-size:16px; color:Red;">    *** TEAM REGISTRATION IS NOW CLOSED. If you are unable to attend the registration party, you can register additional team members at the pool play facility starting 1 hour before the first match in your division. ***
<br />
-->
<br />
</span>
<span style="font-size:14px;">
Please register your team on the <a target="_blank" title="NAGVA website" href="http://nagva.org/tourneys.cfm">NAGVA</a>
 website, then pay your team fees as described below.
 <br />
Please note that registration is based on when your payment is received, not when you submit your roster online on the NAGVA site.
</span>
<br />

<h3>Fees</h3>
<p>
<span style="font-size:14px;">
The fee for one team of 7 is $455 until February 28, 2014.  Additional players (up to 2) are $65 each. After February 28th, the team fee is $505 until March 21st. After March 21st, the team fee is $555 until the final deadline of March 28th.
</span></p>

<h3>Pay with Paypal</h3>

<!--Coming Soon!-->

<!--
<asp:Panel ID="closedPanel" runat="server">
<p>
<span style="font-size:14px; color:red">Registration is closed at this time.</span>
</p>
</asp:Panel>
-->

<asp:Panel ID="openPanel" runat="server">

<!--
JPC 12/30/11
<span style="font-size:14px;">Payments made through PayPal will be charged a $15 fee 
    for the team and $3 for extra players in addition to the normal registration to 
    cover the cost of the transaction.
</span>
-->
    
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
    <ol>
    <li>Enter your team name and select your division (e.g., Team XYZ - BB)</li>
    <li>If your team has more than 7 players, select the number of additional players</li>
    <li>Proceed to PayPal checkout</li>
    <li><b><font color="red">You must click on "Return to CGVA" on the Paypal site once you have made your payment.</font></b></li>
    </ol>
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
Teams will not be available in the drop-down list below to add extra players until the original PayPal payment is processed.  
If you need to add additional players and your team is not visible, please contact the Registrar.
<br /><br />
To add players:
<ol>
<li>Select your team name/division</li>
<li>Select the number of additional players (beyond 7)</li>
<li>Proceed to PayPal checkout</li>
<li><b><font color="red">You must click on "Return to CGVA" on the Paypal site once you have made your payment.</font></b></li>
</ol>
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



<!-- 
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
-->

<!--
<h3>Registration Deadline:</h3>
<span style="font-size:14px;">
<span style="font-size:14px ! important;">
Normal registration closes on March 25th, 2011.  Late registration closes on April 1st, 2011.  
Entries must be postmarked by the given dates to be accepted.  If you are submitting a late registration, 
you MUST email the tournament director to let them know to expect your entry. 

</span></span>
<br />
-->


<h3>Refund Policy</h3>
<span style="font-size:14px;">
<ul>
<li>CGVA will refund 100% of your registration fee if requested on or before March 14, 2013</li>
<li>CGVA will refund 50% of your registration fee if requested between  March 15, 2013 and March 21, 2013</li>
<li>No refunds will be given if requested on or after March 22, 2013</li>
</span>

</div>
</div>

<div id="footer_text">Presented by CGVA, www.cgva.org.  Last Updated 4/1/14</div>
</div>
</div>
</div>
<p><!-- wfxbuild / 1.0 / layout6-15-l1 / 2010-01-07 08:02:55 CET--></p>
       
     
</body></html>