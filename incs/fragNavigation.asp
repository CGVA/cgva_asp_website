<!-- BEGIN:fragNavigation.asp -->
<table width='156' align='left' border="0" cellpadding="1" cellspacing="1" bgcolor="#000066">

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="http://www.cgva.org/index.asp">Home</a></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="http://www.cgva.org/MyCGVA/Login.aspx">MyCGVA</a></td>
</tr>

<!-- Dynamic pull of active events-->
<%For i = 0 to rsActiveEventRows%>
	<tr>
	<td style="background-color:#819FF7; height:25px; text-align:center;">
		<a class='menuWhiteNavigation' 
		href="index.asp?EVENT_CD=<%=rsActiveEventData(0,i)%>">
			<%=rsActiveEventData(1,i)%>
		</a>
	</td>
	</tr>
<%Next%>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=Directions">Directions</a></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=LeagueRules">League Rules</a></span></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=Officiating">Officiating</a></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="/lastdigindenver/">Last Dig</a></td>
</tr>

<!-- JPC 5/2/2009
<tr>
<td><a class='menuWhiteNavigation' href="index.asp?PAGE=Volleypalooza"><img src='../images/volleypalooza_0.gif' border='0' name='volleypalooza' alt='Volleypalooza' id="volleypalooza" onmouseover="imgOver('volleypalooza')" onmouseout="imgOut('volleypalooza')" /></a></td>
</tr>
-->
<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=Mission">About Us</a></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=Photos">Photos</a></td>
</tr>

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=Links">Links</a></td>
</tr>
<!--
<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=MatchMate">Match Mate</a></td>
</tr>
Removed at Ted's request
-->

<tr>
<td style="background-color:#819FF7; height:25px; text-align:center;"><a class='menuWhiteNavigation' href="index.asp?PAGE=ContactUs">Contact Us</td>
</tr>

</table>
<!-- END:fragNavigation.asp -->

