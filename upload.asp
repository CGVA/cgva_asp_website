<!-- #include virtual = "/incs/clsUpload.asp" -->

<%

	class cfs

		dim searchby, searchvalue, MyUploader, count
		dim fso,f

		sub class_initialize()

			if request.totalbytes > 0 then
				set MyUploader = new clsUpload
				deleteDoc = MyUploader.Fields("deleteDoc").Value

				if deleteDoc <> "" then
					set fs = Server.CreateObject("Scripting.FileSystemObject")
					filename = MyUploader.Fields("filename").Value
					foldername = MyUploader.Fields("foldername").Value
					folderPath = request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\"
					filePath = folderPath & filename
					fs.DeleteFile(filePath)
					searchby = MyUploader.Fields("searchby").Value
					searchvalue = MyUploader.Fields("searchvalue").Value

					Session("Err") = "The document was deleted successfully."
					filename = ""
				else
					file1 = MyUploader.Fields("file1").FileName
				end if

			end if

			''print_form

			if file1 <> "" then

				set fs = Server.CreateObject("Scripting.FileSystemObject")
				file1 = MyUploader.Fields("file1").FileName
''				Filepath = MyUploader.Fields("file1").Filepath
				folderPath = request.servervariables("APPL_PHYSICAL_PATH") & "CGVADOCS\"
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
				print_results
			else
				print_results
			end if

		end sub

		sub class_terminate()
		end sub


	sub print_results()
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		set f = fso.GetFolder(Server.Mappath("CGVADOCS/"))
	%>
		<br />

		<div align='center'>
			<font class='cfont8'>Click on a 'Document Name' to view a document, 'Remove' to delete a document, or 'Browse/Upload' to add a document.</font>

		<form name='fileform' encType="multipart/form-data" method="post">
		<input type="file" name="file1" />
		<input type="submit" name="btnsubmit" value="Upload" />
		</form>
		</div>

		<table align="center" bgcolor='#FFFFFF' cellspacing="1" cellpadding="5" border="0">

		<tr bgcolor='#CCCCCC'>
		<th><font class='cfont10'>Document Name</font></th>
		<th><font class='cfont10'>Date Added</font></th>
		</tr>

<%
		for each file in f.files
%>
			<tr bgcolor='#FFFFFF'>
			<td valign='top'><font class='cfont10'><% = file.name %></font></td>
			<td valign='top'><font class='cfont10'><% = file.DateCreated %></font></td>
			</tr>

<%
		next ''file

%>

		</table>

		<br /><br />
<%
 	end sub

end class

set objcfs = new cfs
%>

