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
	dim EVENT_CD,EVENT_CDarray
	dim DIVISION_CD, DIVISION_CDarray
	dim DIVISION_DESC, DIVISION_DESCarray
	dim DIVISION_IMG, DIVISION_IMGarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"EDIT") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Division" then
		EVENT_CD		 	= Replace(Request("EVENT_CD"),"'","''")
		DIVISION_CD		 	= Replace(Request("DIVISION_CD"),"'","''")
		DIVISION_DESC 		= Replace(Request("DIVISION_DESC"),"'","''")
		DIVISION_IMG 		= Replace(Request("DIVISION_IMG"),"'","''")

		'check to see if the location code entered exists'
		checkDivision()

		if exists <> "Y" then
			sql = "INSERT INTO DIVISION_TBL("_
				& "EVENT_CD, "_
				& "DIVISION_CD, "_
				& "DIVISION_DESC, "_
				& "DIVISION_IMG)"_
				& " VALUES("_
				& "'" & EVENT_CD & "', "_
				& "'" & DIVISION_CD & "', "_
				& "'" & DIVISION_DESC & "', "_
				& "'" & DIVISION_IMG & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Division.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Division.asp")


	ElseIf Request("submitChoice") = "Modify Division(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		DIVISION_CD = FixEmptyCell(Request("DIVISION_CD"))
		DIVISION_CDarray = split(DIVISION_CD,", ")

		DIVISION_DESC = FixEmptyCell(Request("DIVISION_DESC"))
		DIVISION_DESCarray = split(DIVISION_DESC,", ")

		DIVISION_IMG = FixEmptyCell(Request("DIVISION_IMG"))
		DIVISION_IMGarray = split(DIVISION_IMG,", ")

		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			''checkDivision()

''			if exists <> "Y" then
				sql = "UPDATE DIVISION_TBL SET "_
					& "EVENT_CD = '" & EVENT_CDarray(i) & "', "_
					& "DIVISION_CD = '" & DIVISION_CDarray(i) & "', "_
					& "DIVISION_DESC = '" & DIVISION_DESCarray(i) & "', "_
					& "DIVISION_IMG = '" & DIVISION_IMGarray(i) & "' "_
					& "WHERE DIVISION_ID = '" & IDarray(i) & "'"

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
		Response.Redirect("Division.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Division.asp")
	End If


'******************************************'

Sub checkDivision()

	If Request("submitChoice") = "Modify Division(s)" then
	''not completed yet
	''	sql = "SELECT DIVISION_CD FROM DIVISION_TBL WHERE DIVISION_CD = '" & UCase(DIVISION_CDarray(i)) & "' "_
	''		& "AND DIVISION_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT DIVISION_CD FROM DIVISION_TBL WHERE DIVISION_CD = '" & DIVISION_CD & "' "_
			& "AND EVENT_CD = '" & EVENT_CD & "'"
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