<!-- BEGIN: fragEventInformation.asp -->
<table>
<tr>
<td>
	<%
	On Error Resume Next
	''Pass the name of the file to the function.
	''Function getFileContents(strIncludeFile)
	  Dim objFSO
	  Dim objText
	  Dim strPage


	  ''Instantiate the FileSystemObject Object.
	  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


	  ''Open the file and pass it to a TextStream Object (objText). The
	  ''"MapPath" function of the Server Object is used to get the
	  ''physical path for the file.
	  strIncludeFile = "incs/" & EVENT_CD & ".asp"
	  ''rw(Server.MapPath(strIncludeFile))
	  		Set objText = objFSO.OpenTextFile(Server.MapPath(strIncludeFile))


	if err.number <> 0 then
		errorCooking = Err.Description
		rw("No event news page has been set up at this time. Please check back later.")
		On Error Goto 0

	else
		''Read and return the contents of the file as a string.
		''getFileContents = objText.ReadAll
		rw(objText.ReadAll)

		objText.Close
		Set objText = Nothing
		Set objFSO = Nothing
		''End Function
	end if

	%>
</td>
</tr>
</table>
<!--END: fragEventInformation.asp -->