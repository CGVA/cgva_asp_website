<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
    'On Error Resume Next
    SQL = "SELECT CONVERT(VARCHAR(10), START_DTG, 101)  as 'START_DTG', "_
            & "CONVERT(VARCHAR(10), END_DTG, 101)  as 'END_DTG', "_
            & "TEAM_FEE_PAYPAL, "_
            & "IND_FEE_PAYPAL, "_
            & "TEAM_FEE_MAIL, "_
            & "IND_FEE_MAIL "_
            & "from LASTDIG_REG_DTG_TBL WHERE getdate() between START_DTG AND END_DTG"
    set rs = cn.Execute(SQL)
	'Response.Write(SQL)
	if not rs.EOF then
		rsActiveRegistrationData = rs.GetRows
		rsActiveRegistrationRows = UBound(rsActiveRegistrationData,2)
	else
		rsActiveRegistrationRows = -1
	end if
	
    SQL = "SELECT DISTINCT(TEAM_NAME + ' - ' + DIVISION) as 'TEAM_NAME' "_
            & "FROM LASTDIG_TEAM_TBL WHERE RECONCILE_DTG IS NOT NULL"
    set rs = cn.Execute(SQL)
	if not rs.EOF then
		rsActiveTeamData = rs.GetRows
		rsActiveTeamRows = UBound(rsActiveTeamData,2)
	else
		rsActiveTeamRows = -1
	end if	
	'Response.Write("HERE")
	'Response.End
 %>
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
<li><a class="main_nav_active_item" href="index.html">Home</a></li>
<li><a class="main_nav_item" href="registration.html">Registration</a></li>
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
 
 <%
 If Session("Err") <> "" Then 
 Response.Write("<span style=""font-size:14px;font-weight:bold;color:#FF0000"">" & Session("Err") & "</span>")
 Session("Err") = ""
 End If
 %>
 
 
 <h3>Registration:</h3>
<span style="font-size:14px;">We are in the process of finalizing our sanctioning with NAGVA.&nbsp; 
Once this is completed you will need to register your team on the 
<a target="_blank" title="NAGVA website" href="http://www.nagva.org">NAGVA</a> website.&nbsp; 
We hope to have sanctioning completed and team registration available by January, 25th, 2010.</span>
<br />

<h3>Fees:</h3>
<p>
<span style="font-size:14px;">Team fee is $420 for up to 7 players, and additional players may be registered for $60 
per player up to the normal registration cut off date.&nbsp; Late registration team fee will be $500, 
and additional players may be registered for $70 per player.</span></p>
<p>
<span style="font-size:14px;">All fees are in U.S. currency and must be paid by either a Cashier's Check or Money Order 
payable to "<span style="font-weight:bold;">CGVA</span>".&nbsp; We are sorry, but Personal Checks will not be accepted.</span>
</p>

<h3>Pay with Paypal</h3>
<%If  rsActiveRegistrationRows = -1 Then%>
    <p>
    <span style="font-size:14px;">Paypal registration is closed at this time.</span>
    </p>
<%ElseIf rsActiveRegistrationRows > -1 Then %>
    <p>
    <span style="font-size:14px;">Paypal is available for a small $15 fee for 6 or 7 member teams and $3 per additional player.</span>
    </p>
    <script type="text/javascript" language="javascript">
        function trim(stringToTrim) {
            return stringToTrim.replace(/^\s+|\s+$/g, "");
        }

        function validateForm1() {
            if (trim(form1.teamName.value) == "")
            {
                alert("Please enter a team name to continue with registration.");
                form1.teamName.focus();
                return false;
            }

            if (form1.division.selectedIndex == 0) {
                alert("Please select your division to continue with registration.");
                form1.division.focus();
                return false;
            }
            return true;
        }

        function validateForm2() {

            if (form2.teamName.selectedIndex == 0) {
                alert("Please select your team name/division to register additional players.");
                form2.teamName.focus();
                return false;
            }

            if (form2.additionalPlayers.selectedIndex == 0) {
                alert("Please select how many additional players you would like to add to your team.");
                form2.additionalPlayers.focus();
                return false;
            }
            return true;
        }
    </script>
    <form id="form1" action="paypal2.asp" onsubmit="return validateForm1();">
    <table border="0" cellspacing='0' cellpadding='3'>
    <tr>
    <td>
        <h4>New Team Registration</h4>
        Enter your team name and select your division* (e.g., Team XYZ - BB)
        <br />
        If your team has more than 7 players, select the number of additional players
        <br />
        Proceed to PayPal checkout
        <p />
    </td>
    </tr>
        
        <tr>
        <td><b>$<%=rsActiveRegistrationData(2,0) %> per team of up to 7 players, if paid by <%=rsActiveRegistrationData(1,0) %></b><br /></td>
        </tr>

        <tr>
        <td>
            <table>
            <tr>
            <td>Team Name <font color="red">*</font></td>
            <td><input type="text"  name="teamName" maxlength="50" size="30" /></td>
            <td>Division <font color="red">*</font></td>
            <td>
                <select name="division">
                <option value="">-SELECT-</option>
                <option value="B">B</option>
                <option value="BB">BB</option>
                <option value="A">Modified A</option>
                </select>
            </td>
            </tr>
            </table>
            <br />
        </td>
        </tr>
        
        <tr>
        <td><b>$<%=rsActiveRegistrationData(3,0)%> per additional player for a team of 8 or more</b><br /></td>
        </tr>

         <tr>
        <td>
            <table>
            <tr>
            <td>Additional Players</td>
            <td colspan="3">
                <select name="additionalPlayers">
                <option value="0">N/A</option>
                <option value="<%=rsActiveRegistrationData(3,0)*1 %>">1 ($<%=rsActiveRegistrationData(3,0) %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*2 %>">2 ($<%=rsActiveRegistrationData(3,0)*2 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*3 %>">3 ($<%=rsActiveRegistrationData(3,0)*3 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*4 %>">4 ($<%=rsActiveRegistrationData(3,0)*4 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*5 %>">5 ($<%=rsActiveRegistrationData(3,0)*5 %>)</option>
                </select>
            </td>
            </tr>
            </table>
            <br />
        </td>
        </tr>

    <tr>
    <td>
    <input type="hidden"  name="registrationType" ID="registrationType"  value="TEAM" />
    <input type="hidden"  name="teamFee" ID="teamFee"  value="<%=rsActiveRegistrationData(2,0) %>" />
    <input type="image"  name="payPalButton" ID="Image1" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" />
    </td>
    </tr>
        
    <tr><td><font color="red">* Required.</font></td></tr>

        </table>

    </form>

    <form id="form2" action="paypal2.asp" onsubmit="return validateForm2();">
    <table border="0" cellspacing='0' cellpadding='3'>
    <tr>
    <td>
    <h4>Existing Team Update</h4>
    Select your team name/division
    <br />
    If your team has more than 7 players, select the number of additional players
    <br />
    Proceed to PayPal checkout
    <p />
    </td>
    </tr>
     <tr>
    <td>
        <table>
        <tr>
                <td>Team Name <font color="red">*</font></td>
        <td>
            <select name="teamName">
            <option value="">-SELECT-</option>
            <%For counter = 0 to rsActiveTeamRows %>
            <option value="<%=rsActiveTeamData(0,counter)%>"><%=rsActiveTeamData(0,counter)%></option>
            <%Next %>
            </select>
        </td>
        <td>Additional Players <font color="red">*</font></td>
        <td>
            <select name="additionalPlayers">
            <option value="0">-SELECT-</option>
                <option value="<%=rsActiveRegistrationData(3,0)*1 %>">1 ($<%=rsActiveRegistrationData(3,0) %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*2 %>">2 ($<%=rsActiveRegistrationData(3,0)*2 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*3 %>">3 ($<%=rsActiveRegistrationData(3,0)*3 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*4 %>">4 ($<%=rsActiveRegistrationData(3,0)*4 %>)</option>
                <option value="<%=rsActiveRegistrationData(3,0)*5 %>">5 ($<%=rsActiveRegistrationData(3,0)*5 %>)</option>
            </select>
        </td>
        </tr>
        </table>
        <br />
    </td>
    </tr>
    
    <tr>
    <td>
    <input type="hidden"  name="registrationType" ID="registrationType2"  value="PLAYER" />
    <input type="image"  name="payPalButton" ID="Image2" src="https://www.paypal.com/en_US/i/btn/btn_xpressCheckout.gif" />
    </td>
    </tr>
        
    <tr><td><font color="red">* Required.</font></td></tr>
    </table>
    </form>
<%End If %>


<h3>Payment by mail</h3>
<p><span style="font-size:14px;">
<span style="font-weight:bold;"><span style="font-size:14px ! important;">Send Payments to:</span></span>
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
<span style="font-size:14px ! important;">Normal registration closes on March 19th, 2010.&nbsp; Late registration closes on 
March, 26th, 2010.&nbsp; Entries must be postmarked by the given dates to be accepted.&nbsp; If you are submitting a 
late registration, you MUST email the tournament director to let them know to expect your entry.
</span></span>
<br />

<h3>Refund Policy:</h3>
CGVA will refund 100% of your registration fee if requested on or before March 12, 2010<br />CGVA will refund 50% of your registration fee if requested between March 13, 2010 and March 19, 2010<br />No refunds will be given if requested on or after March 20, 2010<br /></div>
</div>

<div id="footer_text">Modified 1/6/10</div>
</div>
</div>
</div>
<p><!-- wfxbuild / 1.0 / layout6-15-l1 / 2010-01-07 08:02:55 CET--></p>
       
     
</body></html>