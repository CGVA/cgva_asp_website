<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>PayPal - Error Page</title>
</head>
<body>
    <form id="form1" runat="server">
		<table>
			<tr>
				<td class="field"></td>
				<td><%=Request.QueryString.Get("ErrorCode")%></td>
			</tr>
			
			<tr>
				<td class="field"></td>
				<td><%=Request.QueryString.Get("Desc")%></td>
			</tr>
			
			<tr>
				<td class="field"></td>
				<td><%=Request.QueryString.Get("Desc2")%></td>
			</tr>

		</table>
    </form>
</body>
</html>
