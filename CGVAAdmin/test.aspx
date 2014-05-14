<%@ Page Language="VB" AutoEventWireup="false" CodeFile="test.aspx.vb" Inherits="CGVA_test" %>

<html>
<head>
<title>TinyMCE Test</title>
<script type="text/javascript" src="../jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript">
tinyMCE.init({
	mode : "textareas"
});
</script>
<script type="text/javascript" language="javaScript">
<!--
function validate()
{

	if(FORM.txtEmailFrom.value == "")
	{
		alert("Please enter the 'Email From' address.");
		FORM.txtEmailFrom.focus();
		return false;
	}

	if(FORM.txtEmailTo.value == "" && FORM.txtEmailCC.value == "" && FORM.txtEmailBcc.value == "")
	{
		alert("Please enter the email address(es).");
		FORM.txtEmailTo.focus();
		return false;
	}

	if(FORM.txtEmailSubject.value == "")
	{
		alert("Please enter the 'Email Subject'.");
		FORM.txtEmailSubject.focus();
		return false;
	}

	return true;
}

function ToBcc()
{
	if(FORM.TO_BCC[0].checked)
	{
		FORM.txtEmailBcc.value = FORM.txtEmailTo.value;
		FORM.txtEmailTo.value = "";
	}
	else
	{
		FORM.txtEmailTo.value = FORM.txtEmailBcc.value;
		FORM.txtEmailBcc.value = "";
	}
}



-->
</script>

</head>
<body>
    <form runat="server" enctype="multipart/form-data" method='post'  id='FORM' name='FORM' onsubmit='return validate();'>
<!-- form sends content to moxiecode's demo page -->
 <br />
 <asp:Label ID="txtMsgLabel" runat="server" Visible="False" Width="877px"></asp:Label>
 <br />
<div align="center">
<table border="0" cellspacing="0" cellpadding="4">

<tr align='left' valign='middle'>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial">Email From: </td>
<td><asp:TextBox runat="server" ID="txtEmailFrom" AutoPostBack="true" text="" Columns="50"></asp:TextBox></td>
</tr>

<tr>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email To: </td>
<td valign='top' style="font-size: 10pt; font-family: Arial"><asp:TextBox runat="server" ID="txtEmailTo" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox><input type='radio' name='TO_BCC' value="TO" onclick='ToBcc();' />Move addresses to 'Bcc'</td>
</tr>

<tr>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email CC: </td>
<td><asp:TextBox runat="server" ID="txtEmailCC" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
</tr>

<tr>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Bcc: </td>
<td style="font-size: 10pt; font-family: Arial"><asp:TextBox runat="server" ID="txtEmailBcc" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox><input type='radio' name='TO_BCC' value="BCC" onclick='ToBcc();' />Move addresses to 'To'</td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Subject Line: </td>
<td><asp:TextBox runat="server" ID="txtEmailSubject" Text="" Columns="90" /></td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Attachment 1: </td>
<td>
    <asp:FileUpload ID="FileUpload1" runat="server" Width="566px" /></td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Attachment 2: </td>
<td>
    <asp:FileUpload ID="FileUpload2" runat="server" Width="566px" /></td>
</tr>

<tr>
<td valign="top" align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Body: </td>
<td>
	<textarea name="letter" cols="50" rows="15"></textarea>
</td>
</tr>

<tr>
<td align='right' colspan='2'>
    <br />
    <asp:Button ID="Button1" runat="server" Text="Send Email" />&nbsp;
</td>
</tr>

</table>
</div>
</form>
</body>
</html>
