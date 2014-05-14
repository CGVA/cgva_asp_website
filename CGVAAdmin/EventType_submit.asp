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
	dim ID, EVENT_TYPE_DESC
	dim IDarray, EVENT_TYPE_DESCarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Event Type" then

		EVENT_TYPE_DESC = Replace(Request("EVENT_TYPE_DESC"),"'","''")

		'check to see if the location code entered exists'
		checkEventType()

		if exists <> "Y" then
			sql = "INSERT INTO EVENT_TYPE_TBL("_
				& "EVENT_TYPE_DESC)"_
				& " VALUES("_
				& "'" & EVENT_TYPE_DESC & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("EventType.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("EventType.asp")


	ElseIf Request("submitChoice") = "Modify Event Type(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_TYPE_DESC = FixEmptyCell(Request("EVENT_TYPE_DESC"))
		EVENT_TYPE_DESCarray = split(EVENT_TYPE_DESC,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			''checkLocationCode()

''			if exists <> "Y" then
				sql = "UPDATE EVENT_TYPE_TBL SET "_
					& "EVENT_TYPE_DESC = '" & EVENT_TYPE_DESCarray(i) & "' "_
					& "WHERE EVENT_TYPE_ID = '" & IDarray(i) & "'"

				rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)
''			else
''				notUpdated = notUpdated & " " & IDarray(i)
''			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("EventType.asp?notupdated=" & notUpdated)

	Else
		rw(Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("EventType.asp")
	End If


'******************************************'

Sub checkEventType()

	If Request("submitChoice") = "Modify Event Type(s)" then
		sql = "SELECT EVENT_TYPE_DESC FROM EVENT_TYPE_TBL WHERE UPPER(EVENT_TYPE_DESC) = '" & UCase(EVENT_TYPE_DESCarray(i)) & "' "_
			& "AND ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT EVENT_TYPE_DESC FROM EVENT_TYPE_TBL WHERE UPPER(EVENT_TYPE_DESC) = '" & UCase(EVENT_TYPE_DESC) & "' "
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