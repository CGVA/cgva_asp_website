<!-- #include virtual = "/incs/clsUpload.asp" -->

<%
	Server.ScriptTimeout = 600


	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"BOD") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	if request.totalbytes > 0 then
		set MyUploader = new clsUpload
		deleteDoc = MyUploader.Fields("deleteDoc").Value

		if deleteDoc <> "" then
			set fs = Server.CreateObject("Scripting.FileSystemObject")
			removefilename = MyUploader.Fields("removefilename").Value

			if MyUploader.Fields("removefolderPath").Value="IMAGES" then
				removefolderPath = request.servervariables("APPL_PHYSICAL_PATH") & "images\"
			elseif MyUploader.Fields("removefolderPath").Value="PERSONNEL_IMAGES" then
				removefolderPath = request.servervariables("APPL_PHYSICAL_PATH") & "personnel_images\"
			else
				removefolderPath = request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\" & MyUploader.Fields("removefolderPath").Value & "\"
			end if

			''removefolderPath = request.servervariables("APPL_PHYSICAL_PATH") & MyUploader.Fields("removefolderPath").Value
			filePath = removefolderPath & removefilename
			'rw(filePath)
			''Response.End
			fs.DeleteFile(filePath)

			Session("Err") = "The document was deleted successfully."
			removefilename = ""
		else
			file1 = MyUploader.Fields("file1").FileName
		end if

	end if

	if file1 <> "" then

		set fs = Server.CreateObject("Scripting.FileSystemObject")
		file1 = MyUploader.Fields("file1").FileName
		folder = MyUploader.Fields("folder").Value
''				Filepath = MyUploader.Fields("file1").Filepath
		if folder="IMAGES" then
			folderPath = request.servervariables("APPL_PHYSICAL_PATH") & "images\"
		elseif folder="PERSONNEL_IMAGES" then
			folderPath = request.servervariables("APPL_PHYSICAL_PATH") & "personnel_images\"
		else
			folderPath = request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\" & folder & "\"
		end if

		''folderPath = "\CGVADOCS\"
		filePath = folderpath & file1

''				Response.Write(filePath)
''				Response.End

''				Response.Write(file1)
''				Response.End

''				set fs=Server.CreateObject("Scripting.FileSystemObject")
''				fs.CopyFile Filepath,folderPath
''				set fs=nothing

		MyUploader("file1").SaveAs filePath
		Session("Err") = "The document was uploaded successfully."

	end if

	set fso = Server.CreateObject("Scripting.FileSystemObject")
	set contracts = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\CONTRACTS\")
	set financial = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\FINANCIAL\")
	set general = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\GENERAL\")
	set images = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "IMAGES\")
	set league = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\LEAGUE\")
	set minutes = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\MINUTES\")
	set personnel_images = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "PERSONNEL_IMAGES\")
	set PR = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\PR\")
	set social = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\SOCIAL\")
	set tournament = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\TOURNAMENT\")
	set volunteer = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\VOLUNTEER\")
	set website = fso.GetFolder(request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\WEBSITE\")
%>

<!--#include virtual="/incs/fragHeader.asp"-->

<script language='javascript'>
<!--

	function validateSubmit()
	{

		if(fileform.file1.value == "")
		{
			alert("Please select a document to be uploaded.");
			return false;
		}
		else if(fileform.folder.selectedIndex == 0)
		{
			alert("Please select the folder where the document should be placed.");
			return false;
		}
		else
		{
			return true;
		}

	}

	function validate(foldername,filename)
	{
		//alert(foldername);
		input_box = confirm("Please confirm that you wish to delete this document.");

		if(input_box==true)
		{
			fileform.deleteDoc.value = "Y";
			fileform.removefolderPath.value = foldername;
			fileform.removefilename.value = filename;
			document.fileform.submit();
			//return true;

		}
		else
		{
			//return false;
		}
	}

-->
</script>
</head>

<!-- #include virtual="/incs/rw.asp" -->
<!-- #include virtual="/incs/header.asp" -->
<!-- #include virtual="/incs/fragHeaderGraphics.asp" -->

<tr bgcolor='#FFFFFF'>
<td>
	<%If Instr(Session("ACCESS"),"ADMIN") > 0  Then%>
		<div align='center'>
		<font class='cfont10'>Click on a 'Document Name' to view a document,
		right-click on a 'Document Name' to download it to your computer,
		click 'Remove' to delete a document,
		or click 'Browse/Upload' to add a document.</font>

		<%If Session("Err") <> "" Then
		rw("<br /><br />")
		rw("<font class='cfontSuccess10'>" & Session("Err") & "</font>")
		rw("<br /><br />")
		Session("Err")=""
		End If%>

		<form name='fileform' encType="multipart/form-data" method="post" onSubmit="return validateSubmit();">
		<table>
		<tr>
		<td><font class='cfont8'>File:</td>
		<td><input type="file" name="file1" size='30' /></td>
		</tr>
		<tr>
		<td><font class='cfont8'>Folder:</td>
		<td>
				<select name="folder">
			<OPTION VALUE=''>-select-</OPTION>
			<OPTION VALUE='CONTRACTS'>CONTRACTS</OPTION>
			<OPTION VALUE='FINANCIAL'>FINANCIAL</OPTION>
			<OPTION VALUE='GENERAL'>GENERAL</OPTION>
			<OPTION VALUE='IMAGES'>IMAGES</OPTION>
			<OPTION VALUE='LEAGUE'>LEAGUE</OPTION>
			<OPTION VALUE='MINUTES'>MINUTES</OPTION>
			<OPTION VALUE='PERSONNEL_IMAGES'>PERSONNEL_IMAGES</OPTION>
			<OPTION VALUE='PR'>PR</OPTION>
			<OPTION VALUE='SOCIAL'>SOCIAL</OPTION>
			<OPTION VALUE='TOURNAMENT'>TOURNAMENT</OPTION>
			<OPTION VALUE='VOLUNTEER'>VOLUNTEER</OPTION>
			<OPTION VALUE='WEBSITE'>WEBSITE</OPTION>
		</select>
		</td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td align='right'><input type="submit" name="btnsubmit" value="Upload" /></td>
		</tr>
		</table>
		<input type='hidden' name='deleteDoc' value='' />
		<input type='hidden' name='removefilename' value='' />
		<input type='hidden' name='removefolderPath' value='' />
		</form>
		</div>

		<table align="center" bgcolor='#FFFFFF' cellspacing="1" cellpadding="5" border="0">

		<tr bgcolor='#CCCCCC'>
		<th><font class='cfont10'>Remove</font></th>
		<th><font class='cfont10'>Folder</font></th>
		<th><font class='cfont10'>Document Name</font></th>
		<th><font class='cfont10'>Date Added</font></th>
		</tr>

	<%
		for each file in contracts.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('CONTRACTS','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>CONTRACTS</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/CONTRACTS/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in financial.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('FINANCIAL','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>FINANCIAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/FINANCIAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in general.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href='javascript:validate("GENERAL","<%=file.name%>");'><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>GENERAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/GENERAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in images.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('IMAGES','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>IMAGES</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/images/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in league.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('LEAGUE','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>LEAGUE</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/LEAGUE/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in minutes.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('MINUTES','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>MINUTES</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/MINUTES/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in personnel_images.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('PERSONNEL_IMAGES','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>PERSONNEL IMAGES</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/personnel_images/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in PR.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('PR','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>PR</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/PR/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in social.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('SOCIAL','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>SOCIAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/SOCIAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in tournament.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('TOURNAMENT','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>TOURNAMENT</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/TOURNAMENT/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in volunteer.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('VOLUNTEER','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>VOLUNTEER</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/VOLUNTEER/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in website.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><a class='menuGray' href="javascript:validate('WEBSITE','<%=file.name%>');"><img src='../images/trashcan.jpg' border='0' /></a></font></td>
			<td valign='top'><font class='cfont10'>WEBSITE</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/WEBSITE/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

		<tr height='50'><td>&nbsp;</td></tr>
		</table>

	<%Else%>
		<div align='center'>
		<font class='cfont8'>Click on a 'Document Name' to view a document, right-click on a 'Document Name' to download it to your computer, or click 'Browse/Upload' to add a document.</font>

		<%If Session("Err") <> "" Then
		rw("<br /><br />")
		rw("<font class='cfontSuccess10'>" & Session("Err") & "</font>")
		rw("<br /><br />")
		Session("Err")=""
		End If%>

		<form name='fileform' encType="multipart/form-data" method="post" onSubmit="return validateSubmit();">
		<table>
		<tr>
		<td><font class='cfont8'>File:</td>
		<td><input type="file" name="file1" size='30' /></td>
		</tr>
		<tr>
		<td><font class='cfont8'>Folder:</td>
		<td>
				<select name="folder">
			<OPTION VALUE=''>-select-</OPTION>
			<OPTION VALUE='CONTRACTS'>CONTRACTS</OPTION>
			<OPTION VALUE='FINANCIAL'>FINANCIAL</OPTION>
			<OPTION VALUE='GENERAL'>GENERAL</OPTION>
			<!--<OPTION VALUE='IMAGES'>IMAGES</OPTION>-->
			<OPTION VALUE='LEAGUE'>LEAGUE</OPTION>
			<OPTION VALUE='MINUTES'>MINUTES</OPTION>
			<!--<OPTION VALUE='PERSONNEL_IMAGES'>PERSONNEL_IMAGES</OPTION>-->
			<OPTION VALUE='PR'>PR</OPTION>
			<OPTION VALUE='SOCIAL'>SOCIAL</OPTION>
			<OPTION VALUE='TOURNAMENT'>TOURNAMENT</OPTION>
			<OPTION VALUE='VOLUNTEER'>VOLUNTEER</OPTION>
			<OPTION VALUE='WEBSITE'>WEBSITE</OPTION>
		</select>
		</td>
		</tr>
		<tr>
		<td>&nbsp;</td>
		<td align='right'><input type="submit" name="btnsubmit" value="Upload" /></td>
		</tr>
		</table>

		<input type='hidden' name='deleteDoc' value='' />
		<input type='hidden' name='removefilename' value='' />
		<input type='hidden' name='removefolderPath' value='' />
		</form>
		</div>

		<table align="center" bgcolor='#FFFFFF' cellspacing="1" cellpadding="5" border="0">

		<tr bgcolor='#CCCCCC'>
		<th><font class='cfont10'>Folder</font></th>
		<th><font class='cfont10'>Document Name</font></th>
		<th><font class='cfont10'>Date Added</font></th>
		</tr>


	<%
		for each file in contracts.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>CONTRACTS</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/CONTRACTS/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in financial.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>FINANCIAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/FINANCIAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in general.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>GENERAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/GENERAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>


	<%
		for each file in league.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>LEAGUE</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/LEAGUE/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in minutes.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>MINUTES</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/MINUTES/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>


	<%
		for each file in PR.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>PR</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/PR/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in social.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>SOCIAL</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/SOCIAL/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in tournament.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>TOURNAMENT</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/TOURNAMENT/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in volunteer.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>VOLUNTEER</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/VOLUNTEER/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

	<%
		for each file in website.files
	%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'>WEBSITE</font></td>
			<td valign='top'><font class='cfont10'><a target="_blank" class='menuGray' href='http://cgva.org/CGVADOCS/WEBSITE/<% = file.name %>'><% = file.name %></a></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

	<%
		next ''file
	%>

		<tr height='50'><td>&nbsp;</td></tr>
		</table>

	<%End If%>


</td>
</tr>

<!--#include virtual="/incs/fragContact.asp"-->

</table>

<br /><br />
<br /><br />
<br /><br />

</body>
</html>

