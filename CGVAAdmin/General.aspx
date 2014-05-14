<%@ Page Language="VB"  AutoEventWireup="false" CodeFile="General.aspx.vb" Inherits="General" ValidateRequest="false" Title="Untitled Page" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server"><title>GP Communications Tool</title></head>

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

<body style="margin:0px; padding-right: 0px; padding-left: 0px; padding-bottom: 0px; clip: rect(auto auto auto auto); padding-top: 0px;">

hi

</form>
</body>
</html>