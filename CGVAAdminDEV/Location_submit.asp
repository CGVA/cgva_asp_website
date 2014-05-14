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
	dim ID, LOCATION_CD, LOCATION_DESC, LOCATION_PHONE, LOCATION_FIRST_NAME, LOCATION_LAST_NAME,ADDRESS_LINE1,ADDRESS_LINE2,CITY, STATE, ZIP
	dim IDarray, LOCATION_CDarray, LOCATION_DESCarray, LOCATION_PHONEarray, LOCATION_FIRST_NAMEarray, LOCATION_LAST_NAMEarray
	dim ADDRESS_LINE1array,ADDRESS_LINE2array,CITYarray, STATEarray, ZIParray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Location" then

		LOCATION_CD = Replace(Request("LOCATION_CD"),"'","''")
		LOCATION_DESC = Replace(Request("LOCATION_DESC"),"'","''")
		LOCATION_PHONE = Replace(Request("LOCATION_PHONE"),"'","''")
		LOCATION_FIRST_NAME = Replace(Request("LOCATION_FIRST_NAME"),"'","''")
		LOCATION_LAST_NAME = Replace(Request("LOCATION_LAST_NAME"),"'","''")
		ADDRESS_LINE1 = Replace(Request("ADDRESS_LINE1"),"'","''")
		ADDRESS_LINE2 = Replace(Request("ADDRESS_LINE2"),"'","''")
		CITY = Replace(Request("CITY"),"'","''")
		STATE = Replace(Request("STATE"),"'","''")
		ZIP = Replace(Request("ZIP"),"'","''")

		'check to see if the location code entered exists'
		checkLocationCode()

		if exists <> "Y" then
			sql = "INSERT INTO LOCATION_TBL("_
				& "LOCATION_CD, "_
				& "LOCATION_DESC, "_
				& "LOCATION_PHONE, "_
				& "LOCATION_FIRST_NAME, "_
				& "LOCATION_LAST_NAME, "_
				& "ADDRESS_LINE1, "_
				& "ADDRESS_LINE2, "_
				& "CITY, "_
				& "STATE, "_
				& "ZIP)"_
				& " VALUES("_
				& "'" & LOCATION_CD & "', "_
				& "'" & LOCATION_DESC & "', "_
				& "'" & LOCATION_PHONE & "', "_
				& "'" & LOCATION_FIRST_NAME & "', "_
				& "'" & LOCATION_LAST_NAME & "', "_
				& "'" & ADDRESS_LINE1 & "', "_
				& "'" & ADDRESS_LINE2 & "', "_
				& "'" & CITY & "', "_
				& "'" & STATE & "', "_
				& "'" & ZIP & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Location.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Location.asp")


	ElseIf Request("submitChoice") = "Modify Location(s)" then

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		LOCATION_CD = FixEmptyCell(Request("LOCATION_CD"))
		LOCATION_CDarray = split(LOCATION_CD,", ")

		LOCATION_DESC = FixEmptyCell(Request("LOCATION_DESC"))
		LOCATION_DESCarray = split(LOCATION_DESC,", ")

		LOCATION_PHONE = FixEmptyCell(Request("LOCATION_PHONE"))
		LOCATION_PHONEarray = split(LOCATION_PHONE,", ")

''		rw(UBound(LOCATION_PHONEarray) & "<br />")

		LOCATION_FIRST_NAME = FixEmptyCell(Request("LOCATION_FIRST_NAME"))
		LOCATION_FIRST_NAMEarray = split(LOCATION_FIRST_NAME,", ")

''		rw(UBound(LOCATION_FIRST_NAMEarray) & "<br />")

		LOCATION_LAST_NAME = FixEmptyCell(Request("LOCATION_LAST_NAME"))
		LOCATION_LAST_NAMEarray = split(LOCATION_LAST_NAME,", ")

''		rw(UBound(LOCATION_LAST_NAMEarray) & "<br />")

		ADDRESS_LINE1 = FixEmptyCell(Request("ADDRESS_LINE1"))
		ADDRESS_LINE1array = split(ADDRESS_LINE1,", ")

''		rw(UBound(ADDRESS_LINE1array) & "<br />")

		ADDRESS_LINE2 = FixEmptyCell(Request("ADDRESS_LINE2"))
		ADDRESS_LINE2array = split(ADDRESS_LINE2,", ")

''		rw(UBound(ADDRESS_LINE2array) & "<br />")

		CITY = FixEmptyCell(Request("CITY"))
		CITYarray = split(CITY,", ")

''		rw(UBound(CITYarray) & "<br />")

		STATE = FixEmptyCell(Request("STATE"))
		STATEarray = split(STATE,", ")

''		rw(UBound(STATEarray) & "<br />")

		ZIP = FixEmptyCell(Request("ZIP"))
		ZIParray = split(ZIP,", ")

''		rw(UBound(ZIParray) & "<br /><br />")

		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			'check to see if the location code entered exists if they are updating the location code field'

			if IDarray(i) <> LOCATION_CDarray(i) then
				checkLocationCode()
			else
				exists = "N"
			end if

			if exists <> "Y" then
				sql = "UPDATE LOCATION_TBL SET "_
					& "LOCATION_CD = '" & LOCATION_CDarray(i) & "', "_
					& "LOCATION_DESC = '" & LOCATION_DESCarray(i) & "', "_
					& "LOCATION_PHONE = '" & LOCATION_PHONEarray(i) & "', "_
					& "LOCATION_FIRST_NAME = '" & LOCATION_FIRST_NAMEarray(i) & "', "_
					& "LOCATION_LAST_NAME = '" & LOCATION_LAST_NAMEarray(i) & "', "_
					& "ADDRESS_LINE1 = '" & ADDRESS_LINE1array(i) & "', "_
					& "ADDRESS_LINE2 = '" & ADDRESS_LINE2array(i) & "', "_
					& "CITY = '" & CITYarray(i) & "', "_
					& "STATE = '" & STATEarray(i) & "', "_
					& "ZIP = '" & ZIParray(i) & "' "_
					& "WHERE LOCATION_CD = '" & IDarray(i) & "'"

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
		Response.Redirect("Location.asp?notupdated=" & notUpdated)

	Else
		rw(Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Location.asp")
	End If


'******************************************'

Sub checkLocationCode()

	If Request("submitChoice") = "Modify Location(s)" then
		sql = "SELECT LOCATION_CD FROM LOCATION_TBL "_
			& "WHERE UPPER(LOCATION_CD) = '" & UCase(LOCATION_CDarray(i)) & "'"

	Else
		sql = "SELECT LOCATION_CD FROM LOCATION_TBL WHERE UPPER(LOCATION_CD) = '" & UCase(LOCATION_CD) & "' "
	End If

	''rw(sql)
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