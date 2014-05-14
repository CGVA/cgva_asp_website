<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Home.aspx.vb" Inherits="Home" enableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<title>CGVA - Communications Tool Home Page</title>

<script type="text/javascript" src="jscripts/tiny_mce/tiny_mce.js"></script>
<script type="text/javascript">
<!--
tinyMCE.init({
	mode : "textareas"
});

function addPERSONNELType(e){
	//alert("e value:" + e.value);

	var list = document.getElementById("selvisiblePERSONNELlist");
	var hidden = document.getElementById("visiblePERSONNELtypes");

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

function removePERSONNELType(e){
  e[e.selectedIndex] = null;
	var hidden = document.getElementById("visiblePERSONNELtypes");
	hidden.value = "";
	for (i=0; i<e.length; i++){
			hidden.value += e[i].value + ",";
	}
}

function txtPersonnelChange()
{
    //this.options[this.selectedIndex].firstChild.nodeValue
    //alert(FORM.PERSONNEL.options[i].firstChild.nodeValue);

    sFind = FORM.txtPersonnel.value.toUpperCase();

    if(sFind.length == 0)
    {

		if(FORM.PERSONNEL.options[0].selected == true)
		{
			FORM.PERSONNEL.options[0].selected=true;
		}
		else
		{
			FORM.PERSONNEL.options[0].selected=true;
			FORM.PERSONNEL.options[0].selected=false;
		}

	}
    else
	{
		//alert("here");
		for(i = 0;FORM.PERSONNEL.length - 1;i++)
		{

			try
			{
				if(FORM.PERSONNEL.options[i].firstChild.nodeValue.substring(0,sFind.length).toUpperCase() == sFind)
				{
					//alert("match: " + FORM.PERSONNEL[i].value + ", " + FORM.PERSONNEL.options[i].firstChild.nodeValue);

					//get whether scrolled to item is already selected
					if(FORM.PERSONNEL.options[i].selected == true)
					{
						FORM.PERSONNEL.options[i].selected=true;
					}
					else
					{
						FORM.PERSONNEL.options[i].selected=true;
						FORM.PERSONNEL.options[i].selected=false;
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
	if(FORM.visiblePERSONNELtypes.value == "" && FORM.DISTRO.selectedIndex == -1 && FORM.CUSTOM_DISTRO.value == "")
	{
		alert("Please select a person, previously saved distro, or enter a new custom distro to continue.");
		FORM.PERSONNEL.focus();
		return false;
	}

	//dont chose both of these
	if(FORM.DISTRO.selectedIndex > 0 && FORM.CUSTOM_DISTRO.value != "")
	{
		alert("Please select either a previously saved distro or enter a new custom distro to continue (not both).");
		FORM.DISTRO.focus();
		return false;
	}

	//dont chose both of these
	if((FORM.visiblePERSONNELtypes.value != "" && FORM.CUSTOM_DISTRO.value != "")|| (FORM.ADD_ADDRESS.value != "" && FORM.CUSTOM_DISTRO.value != ""))
	{
		alert("Please select either a person/additional address(es) or enter a new custom distro to continue (not both).");
		FORM.DISTRO.focus();
		return false;
	}

	//make sure name for distro is entered if checkbox checked
	//alert(FORM.SAVEDISTRO.checked);
	//alert(trim(FORM.DISTRO_NAME.value).length);
	if(FORM.SAVEDISTRO.checked==1 && trim(FORM.DISTRO_NAME.value).length==0)
	{
		alert("Please enter a distro name to save the email addresses as a distro, or uncheck the box to save this list as a distro.");
		FORM.SAVEDISTRO.focus();
		return false;
	}
	
	return true;

}

function PostPage()
{
    //document.FORM.action = "test.aspx";
    //document.FORM.method = "POST";
    //document.FORM.submit();
}

function checkChange() 
{
    if(FORM.SAVEDISTRO.checked==1)
     {
        FORM.DISTRO_NAME.disabled=false;
     }
     else
     {
        FORM.DISTRO_NAME.value = "";
        FORM.DISTRO_NAME.disabled=true;
     }
}

function trim(stringToTrim) 
{ 
    return stringToTrim.replace(/^\s+|\s+$/g,""); 
} 
-->
</script>

</head>
<link rel="stylesheet" type="text/css" href="incs/style.css" />
<body>

    <form id="FORM" name="FORM" runat="server" onsubmit="return validate();">
    <div align='center' style="background-color:white">
    <b><u><asp:Label class="cfont14" ID="Label1" runat="server" Text="CGVA - Communications Tool"></asp:Label></u></b>
        <br />
        <table align="center" style="width:90%;background-color:White">
            <tr>
                <td align="left" class="cfont10">
                    <asp:BulletedList ID="BulletedList1" style="left:auto" runat="server">
                        <asp:ListItem>Select a person from the CGVA personnel, a saved distro or create your own custom distribution list.</asp:ListItem>
                    	<asp:ListItem>To assist with personnel selection, you may start entering a name in the small textbox. When the name you want to add appears, click on the name to add it to the distro list. To remove a name from the distro, double-click the name in the select box on the right.</asp:ListItem>
	                    <asp:ListItem>If you would like to add additional addresses to the selected personnel, type them in the available textbox, separating each address with a semicolon.</asp:ListItem>
	                    <asp:ListItem>Check the 'Save Distro?' checkbox to save the selected personnel or custom distro for future use.</asp:ListItem>
	                    <asp:ListItem> Then click 'Next ->'.</asp:ListItem>
                    </asp:BulletedList>
                </td>
            </tr>
        </table>
    <asp:Label class="cfonterror10" ID="errorLabel" runat="server" Visible="False"></asp:Label>
    
    <table style="background-color:White" align='center' cellpadding='3' cellspacing='0' border='0'>

    <tr>
    <td colspan='2'>&nbsp;</td>
    </tr>

    <tr>
    <td class="cfont10" valign='top' align='right'><b>CGVA Personnel:</b></td>
    <td align="left">
	    <input type='text' name='txtPersonnel' size='6' onKeyUp="txtPersonnelChange();" />
	    <br />
	    <asp:ListBox size='6' name='PERSONNEL' ID="PERSONNEL" 
            runat="server" onClick="addPERSONNELType(this);" 
            style="margin-top: 0px" Rows="6"></asp:ListBox>
	    &nbsp;
	    <select style="width:350px;" size='6' name='selvisiblePERSONNELlist' ondblclick="removePERSONNELType(this);"></select>
	    <input type="hidden" name="visiblePERSONNELtypes" id="visiblePERSONNELtypes" />
    </td>
    </tr>

    <tr>
    <td class="cfont10" valign='top' align='right'><b>Additional address(es):</b></td>
    <td align="left">
	    <input type='text' name='ADD_ADDRESS' ID='ADD_ADDRESS' size='100' maxlength='500' />
    </td>
    </tr>

    <tr>
    <td class="cfont10" valign='top' align='right'><b>Custom Distro:</b></td>
    <td align="left">
	    <input type='text' name='CUSTOM_DISTRO' ID='CUSTOM_DISTRO' size='50' maxlength='500' />
    </td>
    </tr>

    <tr>
    <td class="cfont10" valign='top' align='right'><b>Save As New Distro?</b></td>
    <td align="left">
	    <!--<asp:CheckBox runat="server" ID="SAVEDISTRO" name='SAVEDISTRO' onchange="checkChange();" />-->
	    <input type="checkbox" id="SAVEDISTRO" name="SAVEDISTRO" onclick="checkChange();" />
	    &nbsp;
	    <!--<asp:TextBox runat="server" ID="DISTRO_NAME" name="DISTRO_NAME" visible="false" />-->
	    <input type='text' name='DISTRO_NAME' id='DISTRO_NAME' disabled="disabled" />
    </td>
    </tr>

    <tr>
    <td class="cfont10" valign='top' align='right'><b>Previously Saved Distros:</b></td>
    <td align="left">
	    <asp:ListBox size='6' name='DISTRO' ID="DISTRO" 
            runat="server" Rows="6"></asp:ListBox>
    </td>
    </tr>

    <tr>
    <td colspan='2' align='right'><asp:Button runat="server" ID="submitButton" type="submit" text="Next ->"  /></td>
    </tr>

    </table>
    </div>

</form>
</body>
</html>
