<%@ Language=VBScript %>

<!-- #include virtual = "/incs/dbConnection.inc" -->

<%
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("index.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If


	''clear session variables
	Session("MGMT") = ""
	Session("MGMT_DATA") = ""
	Session("DISTRO") = ""
	Session("CUSTOM_DISTRO") = ""
	Session("SAVEDISTRO") = ""
	Session("emailFrom") = ""
	Session("emailTo") = ""
	Session("emailSubject") = ""
	Session("emailCC") = ""
	Session("emailBCC") = ""
	Session("ADD_ADDRESS") = ""
	Session("File1") = ""
	Session("File2") = ""


	SQL = "SELECT LAST_NAME,FIRST_NAME " & _
		"FROM db_accessadmin.PERSON_TBL " & _
		"ORDER BY LAST_NAME,FIRST_NAME"
	rw("<!-- SQL: " & SQL & " -->")
	set rs = cn.Execute(SQL)

	if not rs.EOF then
		rsNameData = rs.GetRows
		rsNameRows = UBound(rsNameData,2)
	else
		rsNameRows = -1
	end if

'	SQL = "SELECT ID,TYPE_DESC,[NAME] FROM T_EMAIL_DISTRO a " & _
'		"JOIN T_EMAIL_TYPE b ON a.TYPE = b.TYPE " & _
'		"ORDER BY TYPE_DESC,[NAME]"
'
'	rw("<!-- SQL: " & SQL & " -->")
'	set rs = cn.Execute(SQL)
'
'	if not rs.EOF then
'		rsDistroData = rs.GetRows
'		rsDistroRows = UBound(rsDistroData,2)
'	else
'		rsDistroRows = -1
'	end if

	rs.Close
	set rs = Nothing
	cn.Close
	set cn = Nothing

%>

<!-- #include virtual="/incs/fragHeader.asp" -->

<title>CGVA - Communication Tool Home Page</title>

<script type="text/javascript" src="../jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript">
<!--
tinyMCE.init({
	mode : "textareas"
});

function addMGMTType(e){
	//alert("e value:" + e.value);

	var list = document.getElementById("selvisibleMGMTlist");
	var hidden = document.getElementById("visibleMGMTtypes");

	//alert("list length:" + list.length);
	for (i=0; i<list.length; i++){
			if (e.value == list[i].value) return;
	}

	var opt = new Option(e[e.selectedIndex].text, e.value);
	list.options.add(opt);
	//alert("after list add");
	hidden.value += e.value + ",";
	e.value="";
	//alert("HV:" + hidden.value);
}

function removeMGMTType(e){
  e[e.selectedIndex] = null;
	var hidden = document.getElementById("visibleMGMTtypes");
	hidden.value = "";
	for (i=0; i<e.length; i++){
			hidden.value += e[i].value + ",";
	}
}

function txtMGMTChange()
{
    //this.options[this.selectedIndex].firstChild.nodeValue
    //alert(FORM.MGMT.options[i].firstChild.nodeValue);

    sFind = FORM.txtMGMT.value.toUpperCase();

    if(sFind.length == 0)
    {

		if(FORM.MGMT.options[0].selected == true)
		{
			FORM.MGMT.options[0].selected=true;
		}
		else
		{
			FORM.MGMT.options[0].selected=true;
			FORM.MGMT.options[0].selected=false;
		}

	}
    else
	{
		//alert("here");
		for(i = 0;FORM.MGMT.length - 1;i++)
		{

			try
			{
				if(FORM.MGMT.options[i].firstChild.nodeValue.substring(0,sFind.length).toUpperCase() == sFind)
				{
					//alert("match: " + FORM.MGMT[i].value + ", " + FORM.MGMT.options[i].firstChild.nodeValue);

					//get whether scrolled to item is already selected
					if(FORM.MGMT.options[i].selected == true)
					{
						FORM.MGMT.options[i].selected=true;
					}
					else
					{
						FORM.MGMT.options[i].selected=true;
						FORM.MGMT.options[i].selected=false;
					}

					break;
				}

			}
			catch(e)
			{
				break;
			}

		}

    }
}

function GetSelectedItem(field)
{

	len = eval("FORM." + field + ".length")
	i = 0
	chosen = ""

	for (i = 0; i < len; i++)
	{
		if (eval("FORM." + field + "[i].selected"))
		{
			chosen = chosen + "YES";
		}
	}

	return chosen;
}

function validate()
{

	//alert(FORM.DISTRO.selectedIndex);
	if(FORM.visibleMGMTtypes.value == "" && FORM.DISTRO.selectedIndex == -1 && FORM.CUSTOM_DISTRO.value == "")
	{
		alert("Please select a management value, saved distro or enter a custom distro to continue.");
		FORM.MGMT.focus();
		return false;
	}

	//dont chose both of these
	if(FORM.DISTRO.selectedIndex > 0 && FORM.CUSTOM_DISTRO.value != "")
	{
		alert("Please select either a saved distro or enter a custom distro to continue (not both).");
		FORM.DISTRO.focus();
		return false;
	}

	//dont chose both of these
	if((FORM.visibleMGMTtypes.value != "" && FORM.CUSTOM_DISTRO.value != "")|| (FORM.ADD_ADDRESS.value != "" && FORM.CUSTOM_DISTRO.value != ""))
	{
		alert("Please select either a mgmt list/add address(es) or enter a custom distro to continue (not both).");
		FORM.DISTRO.focus();
		return false;
	}
	return true;

}
-->
</script>

</head>
<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->

<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<div align='center'>
<font class='cfont14'><b><u>CGVA - Communication Tool</u></b></font>
</div>
<br />

<tr bgcolor='#FFFFFF'>
<td>
    <table width='90%' align='center'>

    <tr>
    <td>
	    <ul>
	    <li><font class='cfont10'>Select a value from management, a saved distro or create your own custom distribution list.</font></li>
	    <li><font class='cfont10'>To assist with management selection, you may start entering a name in the small textbox. When the name you want to add appears, click on the name to add it to the distro list. To remove a name from the distro, double-click the name in the select box on the right.</font></li>
	    <li><font class='cfont10'>If you would like to add additional addresses to the selected management, type them in the available textbox, separating each address with a semicolon.</font></li>
	    <li><font class='cfont10'>Check the 'Save Distro?' checkbox to save the selected management or custom distro for future use.</font></li>
	    <li><font class='cfont10'> Then click 'Next ->'.</font></li>
	    </ul>
    </td>
    </tr>

    </table>

    <br /><br />

    <%
	    If Session("Err") <> "" then
		    rw("<div align='center'><font class='cfontSuccess10'><b>" & Session("Err") & "</b></font></div>")
		    rw("<br /><br />")
	    End If

	    Session("Err") = ""
    %>

    <form method='post' action='Email.asp' name='FORM' onSubmit='return validate();'>
    <table align='center' cellpadding='3' cellspacing='0' border='0'>

    <tr>
    <td colspan='2'><font class='cfont10'>&nbsp;</font></td>
    </tr>

    <tr>
    <td valign='top' align='right'><font class='cfont10'><b>Management:</b></font></td>
    <td>
	    <input type='text' name='txtMGMT' size='6' onKeyUp="txtMGMTChange();" />
	    <br />
	    <select name='MGMT' size='6' onClick="addMGMTType(this);">

	    <%
		    For i = 0 to rsNameRows
			    rw("<option value='" & rsNameData(0,i) & "'>" & rsNameData(1,i) & ", " & rsNameData(2,i) & "</option>")
		    Next
	    %>
	    </select>
	    &nbsp;
	    <select style="width:350px;" size='6' name='selvisibleMGMTlist' ondblclick="removeMGMTType(this);"></select>
	    <input type="hidden" name="visibleMGMTtypes" id="visibleMGMTtypes" >
    </td>
    </tr>

    <tr>
    <td valign='top' align='right'><font class='cfont10'><b>Additional address(es):</b></font></td>
    <td>
	    <input type='text' name='ADD_ADDRESS' size='100' maxlength='500' />
    </td>
    </tr>

    <tr>
    <td valign='top' align='right'><font class='cfont10'><b>Custom Distro:</b></font></td>
    <td>
	    <input type='text' name='CUSTOM_DISTRO' size='50' maxlength='500' />
    </td>
    </tr>

    <tr>
    <td valign='top' align='right'><font class='cfont10'><b>Save Distro?</b></font></td>
    <td>
	    <input type='checkbox' name='SAVEDISTRO' />
    </td>
    </tr>



    <tr>
    <td valign='top' align='right'><font class='cfont10'><b>Saved Distros:</b></font></td>
    <td>
	    <select name='DISTRO' size='5' multiple>
	    <!--<option value=''>- Select -</option>-->

	    <%
		    'For i = 0 to rsDistroRows
		    '	rw("<option value='" & rsDistroData(0,i) & "'>" & rsDistroData(1,i) & "-" & rsDistroData(2,i) & "</option>")
		    'Next
	    %>
	    </select>
    </td>
    </tr>

    <tr>
    <td colspan='2' align='right'><input type='submit' value='Next ->' /></td>
    </tr>

    </table>

    </form>
</td>
</tr>
<br />

<script language='Javascript'>
	FORM.MGMT.focus();
</script>

<!--#include virtual="/incs/fragContact.asp"-->

</table>

<%


'******************************************'

Sub closePage()
	rw("</body>")
	rw("</html>")
End Sub

'******************************************'

%>
