<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MyEvents3.aspx.vb" Inherits="MyEvents3" enableEventValidation="false" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>MyCGVA - My Events - Waiver Agreement</title>

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
	<td height="20" class="cfont16">MyCGVA - My Events - Waiver Agreement</td>
	</tr>

	<tr>
	<td height='20'></td>
	</tr>

    <tr bgcolor="#FFFFFF">
    <td>

		<div align='center'>
		<font class='cfont10'><b>
		Please read the following information. If you agree, electronically sign the waiver to 
		continue with registration.
		</b>
		<br />
        </div>

        <form id="Form1" name="INSERT" runat="server">
        <table cellspacing='0' align='center' cellpadding='3'>

        <tr>
        <td align='center'>
        <font class="cfontError10">
        <asp:Label id="messageLabel" runat="server" Text=""></asp:Label>
        </font>
        </td>
        </tr>
        
        <tr>
        <td>
        <asp:Label id="LeagueWaiver" runat="server" Text="
        <font class='cfont9'>
            RELEASE OF LIABILITY AND AUTHORIZATION FOR EVENT PROMOTION
            <p />
            This release and authorization is for the CGVA Fall League occurring at DIVE, 
            from August 25, 2014 through November 17, 2014.
            <p />
            KNOW ALL PERSONS BY THESE PRESENTS that the undersigned, intending to be legally bound, and 
            being of lawful age, do/does hereby for myself, and for my heirs,
            executors, administrators, successors and/or assigns agree to release, remise, acquit, 
            hold harmless, and forever discharge the Colorado Gay Volleyball Association (CGVA) 
            and its agents, officers, directors, board members, other members, tournament directors, 
            league directors, sponsors, event sites, supporters, servants, successors, heirs, executors, 
            or other associates or affiliates, and all other persons, firms corporations, or partnerships 
            acting or related to CGVA or the above event, for and from any and all claims, causes of 
            actions, demands, rights, damages, including medical expenses, pain, suffering, loss of 
            income or work, services, expenses, or other claim related to property loss, or personal 
            injury regarding any participation by the undersigned in the above referenced function. 
            The undersigned also understands that any and all participation is performed solely at my 
            own risk and is not the responsibility of CGVA, or its agents, officers, directors, board 
            members, other members, tournament directors, league directors, sponsors, event sites, 
            supporters, servants, successors, heirs, executors, or other associates or affiliates, 
            and all other persons, firms corporations, or partnerships acting or related to CGVA. 
            The undersigned also certifies that he/she is physically fit, and had not otherwise been 
            informed or advised by a physician or other health care provider not to participate in 
            activities such as the above named event. I acknowledge that I am aware of the risks 
            inherent in volleyball, including possible permanent disability or death and agree to 
            assume all of those risks. I further agree to abide by and be governed by the rules and 
            regulations for this event.
            <p />
            Authorization for event promotion: The undersigned agrees to be filmed, photographed, taped, 
            quoted, or otherwise mentioned, without compensation, by CGVA or anyone
            authorized by CGVA. This includes but is not limited to any official or other authorized 
            photographer, writer, host, or sponsor of the above referenced event, under the conditions 
            authorized by CGVA. The undersigned hereby gives CGVA, and anyone authorized by CGVA the right 
            to use, without compensation, my name, picture, likeness, quotes, and biographical information, 
            whether audio or visual, before, during, and after the period of the undersigned person’s 
            individual or team participation in the above referenced event.  
            <p />
            The undersigned hereby acknowledges having read the entire foregoing Release of Liability and 
            Authorization for Event Promotion, and that the undersigned understands
            and agrees to its contents without exception. The undersigned further represents that he/she is 
            over the age of 18 years and competent to execute this document.
            </font>
            ">
            </asp:Label>
        </td>
        </tr>
        
        <tr>
		<td align="right"><font class='cfont10'>
            <asp:CheckBox id="verifyWaiverInfo" text="I have read and accept the waiver information." runat="server" />
            </font>
            &nbsp;<asp:Button id="submitButton" runat="server" Text="Continue" />
            </td>
		</tr>

		</table>

		</form>

</td>
</tr>

<!-- #include virtual="/incs/fragSiteContact.asp" -->

</table>

</body>
</html>



