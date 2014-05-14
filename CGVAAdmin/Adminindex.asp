<%Session("ACCESS")=""%>

<!--#include virtual="/incs/fragHeader.asp"-->
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphicsTEST.asp" -->

<tr bgcolor="#FFFFFF">
<td>
	<%
	If Session("Err") <> "" Then
		rw("<div align='center'><font class='cfontError10'>" & Session("Err") & "</font></div>")
		Session("Err") = ""
	End If
	%>

	<form name='form' action='Adminindex_submit.asp'>
	<table align='center' cellpadding='3'>
	<tr>
	<td><font class='cfont10'>USER NAME</font></td>
	<td><font class='cfont10'><INPUT TYPE='TEXT' MAXLENGTH='50' NAME='USERNAME' /></font></td>
	</tr>
	<tr>
	<td><font class='cfont10'>PASSWORD</font></td>
	<td><font class='cfont10'><INPUT TYPE='PASSWORD' MAXLENGTH='50' NAME='PASSWORD' /></font></td>
	</tr>
	<tr>
	<td colspan='2'><font class='cfont10'><input type='submit' name='SUBMIT' value='Enter' /></font></td>
	</tr>
	</table>
	</form>
</td>
</tr>

</table>
<script language='Javascript'>
<!--
	form.USERNAME.focus();
//-->
</script>
</body>
</html>
