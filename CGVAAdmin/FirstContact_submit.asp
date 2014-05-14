<%@ Language=VBScript %>
<!-- #include virtual = "/incs/dbConnection.inc" -->
<!-- #include virtual="/incs/rw.asp" -->

<%
	''On Error Resume Next
	Response.Expires = -1
	Response.Buffer = true
	Response.Clear
	Response.CacheControl = "no-cache"
	Server.ScriptTimeout = 600

	dim exists
	dim ID, IDarray
	dim FIRST_CONTACT_DESC,FIRST_CONTACT_DESCarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add First Contact" then
		FIRST_CONTACT_DESC 		= Replace(Request("FIRST_CONTACT_DESC"),"'","''")

		'check to see if the location code entered exists'
		checkFirstContact()

		if exists <> "Y" then
			sql = "INSERT INTO FIRST_CONTACT_TBL("_
				& "FIRST_CONTACT_DESC)"_
				& " VALUES("_
				& "'" & FIRST_CONTACT_DESC & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("FirstContact.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("FirstContact.asp")


	ElseIf Request("submitChoice") = "Modify First Contact(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		FIRST_CONTACT_DESC = FixEmptyCell(Request("FIRST_CONTACT_DESC"))
		FIRST_CONTACT_DESCarray = split(FIRST_CONTACT_DESC,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			checkFirstContact()

			if exists <> "Y" then
				sql = "UPDATE FIRST_CONTACT_TBL SET "_
					& "FIRST_CONTACT_DESC = '" & FIRST_CONTACT_DESCarray(i) & "' "_
					& "WHERE FIRST_CONTACT_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " " & IDarray(i)
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("FirstContact.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("FirstContact.asp")
	End If


'******************************************'

Sub checkFirstContact()

	If Request("submitChoice") = "Modify First Contact(s)" then
		sql = "SELECT FIRST_CONTACT_DESC FROM FIRST_CONTACT_TBL WHERE UPPER(FIRST_CONTACT_DESC) = '" & UCase(FIRST_CONTACT_DESCarray(i)) & "' "_
			& "AND FIRST_CONTACT_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT FIRST_CONTACT_DESC FROM FIRST_CONTACT_TBL WHERE UPPER(FIRST_CONTACT_DESC) = '" & UCase(FIRST_CONTACT_DESC) & "' "
	End If

	rw(sql)
	''Response.End

	set rs = cn.Execute(sql)

	If not rs.EOF then
		exists = "Y"
	Else
		exists = "N"
	End If

End Sub

'******************************************'

Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function
%>