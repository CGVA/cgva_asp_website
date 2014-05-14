<!--#include file="fragHeader.asp"-->
<body cellpadding='0' cellspacing='0'>
<!--#include file="fragHeaderGraphics.asp"-->

<tr bgcolor="#FFFFFF">
<td>
	<form name='admin'>
	Select an area to update:
	<select name='choice' onChange="changePage();">
	<option value='admin.asp'>-select-</option>
	<option value='Division.asp'>Divisions</option>
	<option value='Event.asp'>Events</option>
	<option value='EventType.asp'>Event Types</option>
	<option value='FirstContact.asp'>First Contacts</option>
	<option value='Location.asp'>Locations</option>
	<option value='Person.asp'>Personnel</option>
	</select>
	</form>
</td>
</tr>


<!--#include file="fragContact.asp"-->

</table>
<script language='Javascript'>
<!--
	function changePage()
	{
		choice = admin.choice.options[admin.choice.selectedIndex].value;
		//alert(choice);
		window.location = choice;
	}
//-->
</script>
</body>
</html>
