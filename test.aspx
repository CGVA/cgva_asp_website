<%@ Page Language="VB" AutoEventWireup="false" CodeFile="test.aspx.vb" Inherits="test" validateRequest="false" enableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server"><title>CGVA Communications Tool</title>
    <script type="text/javascript" src="jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript"> 
tinyMCE.init({ 
// General options 
    mode : "specific_textareas",
	editor_selector : "mceEditor",
 
theme : "advanced", 
plugins : "safari,spellchecker,pagebreak,style,layer,table,save,advhr,advimage,advlink,emotions,iespell,inlinepopups,insertdatetime,preview,media,searchreplace,print,contextmenu,paste,directionality,fullscreen,noneditable,visualchars,nonbreaking,xhtmlxtras,template,imagemanager,filemanager", 
remove_script_host : false,
relative_urls : false,
// Theme options 
theme_advanced_buttons1 : "save,newdocument,|,bold,italic,underline,strikethrough,|,justifyleft,justifycenter,justifyright,justifyfull,|,styleselect,formatselect,fontselect,fontsizeselect", 
theme_advanced_buttons2 : "cut,copy,paste,pastetext,pasteword,|,search,replace,|,bullist,numlist,|,outdent,indent,blockquote,|,undo,redo,|,link,unlink,anchor,image,cleanup,help,code,|,insertdate,inserttime,preview,|,forecolor,backcolor", 
theme_advanced_buttons3 : "tablecontrols,|,hr,removeformat,visualaid,|,sub,sup,|,charmap,emotions,iespell,media,advhr,|,print,|,ltr,rtl,|,fullscreen", 
theme_advanced_buttons4 : "insertlayer,moveforward,movebackward,absolute,|,styleprops,spellchecker,|,cite,abbr,acronym,del,ins,attribs,|,visualchars,nonbreaking,template,blockquote,pagebreak,|,insertfile,insertimage", 
theme_advanced_toolbar_location : "top", 
theme_advanced_toolbar_align : "left", 
theme_advanced_statusbar_location : "bottom", 
theme_advanced_resizing : true 
 
// Example content CSS (should be your site CSS) 
//content_css : "css/example.css", 
 
// Drop lists for link/image/media/template dialogs 
//template_external_list_url : "js/template_list.js", 
//external_link_list_url : "js/link_list.js", 
//external_image_list_url : "js/image_list.js", 
//media_external_list_url : "js/media_list.js", 
 
// Replace values for the template plugin 
//template_replace_values : { 
//username : "Some User", 
//staffid : "991234" 
//} 
}); 
</script>




<script type="text/javascript" language="javaScript">
<!--
function validate()
{

	if(EMAIL_FORM.txtEmailFrom.value == "")
	{
		alert("Please enter the 'Email From' address.");
		EMAIL_FORM.txtEmailFrom.focus();
		return false;
	}

	if(EMAIL_FORM.txtEmailTo.value == "" && EMAIL_FORM.txtEmailCC.value == "" && EMAIL_FORM.txtEmailBcc.value == "")
	{
		alert("Please enter the email address(es).");
		EMAIL_FORM.txtEmailTo.focus();
		return false;
	}

	if(EMAIL_FORM.txtEmailSubject.value == "")
	{
		alert("Please enter the 'Email Subject'.");
		EMAIL_FORM.txtEmailSubject.focus();
		return false;
	}

	return true;
}

function ToBcc()
{
	if(EMAIL_FORM.TO_BCC[0].checked)
	{
		EMAIL_FORM.txtEmailBcc.value = EMAIL_FORM.txtEmailTo.value;
		EMAIL_FORM.txtEmailTo.value = "";
	}
	else
	{
		EMAIL_FORM.txtEmailTo.value = EMAIL_FORM.txtEmailBcc.value;
		EMAIL_FORM.txtEmailBcc.value = "";
	}
}



-->
</script>

</head>
<body>
    <form runat="server" enctype="multipart/form-data" method='post'  id='EMAIL_FORM' name='EMAIL_FORM' onsubmit='return validate();'>
 <br />
 <div align='center'>
 <asp:Label ID="txtMsgLabel" runat="server" Visible="False" Width="877px"></asp:Label>
 </div>
 <br />
<div align="center">
<table width='100%' border="0" cellspacing="0" cellpadding="4">

<tr align='left' valign='middle'>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial">Email From: </td>
<td><asp:TextBox runat="server" ID="txtEmailFrom" AutoPostBack="true" text="" Columns="50"></asp:TextBox></td>
</tr>

<tr align='left'>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email To: </td>
<td valign='top' style="font-size: 10pt; font-family: Arial"><asp:TextBox runat="server" ID="txtEmailTo" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox><input type='radio' name='TO_BCC' value="TO" onclick='ToBcc();' />Move addresses to 'Bcc'</td>
</tr>

<tr align='left'>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email CC: </td>
<td valign='top' style="font-size: 10pt; font-family: Arial"><asp:TextBox runat="server" ID="txtEmailCC" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
</tr>

<tr align='left'>
<td valign='top' align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Bcc: </td>
<td valign='top' style="font-size: 10pt; font-family: Arial"><asp:TextBox runat="server" ID="txtEmailBcc" text="" Columns="70" Rows="7" TextMode="MultiLine"></asp:TextBox><input type='radio' name='TO_BCC' value="BCC" onclick='ToBcc();' />Move addresses to 'To'</td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Subject Line: </td>
<td align='left'><asp:TextBox runat="server" ID="txtEmailSubject" Text="" Columns="90" /></td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Attachment 1: </td>
<td align='left'>
    <asp:FileUpload ID="FileUpload1" runat="server" Width="566px" /></td>
</tr>

<tr>
<td align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Attachment 2: </td>
<td align='left'>
    <asp:FileUpload ID="FileUpload2" runat="server" Width="566px" /></td>
</tr>

<tr>
<td valign="top" align='right' style="font-weight: bold; font-size: 10pt; font-family: Arial; height: 32px">Email Body: </td>
<td align="left">
	<!--<textarea id='letter' name="letter" cols="50" rows="15" class="mceEditor"></textarea>-->
<asp:TextBox ID="letter" runat="server" Rows="10" TextMode="MultiLine" Width="557px" Height="267px" class="mceEditor"></asp:TextBox>

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
