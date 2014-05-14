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

	dim vFILTER
	dim exists
	dim ID, IDarray
	dim PERSON_ID, PERSON_IDarray
	dim EFF_DATE, EFF_DATEarray
	dim END_DATE, END_DATEarray
	dim CERT_ID, CERT_IDarray
	dim NOTES, NOTESarray


	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If


	''rw("here")
	''Response.End

	vFILTER = Request("FILTER")

	'determine which record to modify based on the ID number of the record'
	ID						= Request("ID")
	IDarray 				= split(ID,", ")

	EFF_DATE   				= FixEmptyCell(Replace(Request("EFF_DATE"),"'","''"))
	EFF_DATEarray			= split(EFF_DATE,", ")

	END_DATE   				= FixEmptyCell(Replace(Request("END_DATE"),"'","''"))
	END_DATEarray			= split(END_DATE,", ")

	CERT_ID   			= FixEmptyCell(Replace(Request("CERT_ID"),"'","''"))
	CERT_IDarray		= split(CERT_ID,", ")

	NOTES					= FixEmptyCell(Replace(Request("NOTES"),"'","''"))
	NOTESarray				= split(NOTES,", ")



	if Request("submitChoice") <> "" then
		''check for valid dates
		For i = 0 to UBound(IDarray)
			''rw(END_DATEarray(i) & "-" & EFF_DATEarray(i) & "<br />")
			''If END_DATEarray(i) <= EFF_DATEarray(i) then
			''	rw("<=")
			''Else
			''	rw(">")
			''End If

			If  EFF_DATEarray(i) <> "" or END_DATEarray(i) <> "" Then

				If not (IsDate(EFF_DATEarray(i)) or IsDate(END_DATEarray(i))) then
			        ''go back to prior page with error message
			        Session("Err") = "*** ERROR: No updates/entries have been made. Make sure all records have valid dates for effective/end dates<br />and end dates are later than effective dates. ***"
			        Response.Redirect("PlayerRefCertification.asp?FILTER=" & vFILTER & "&notupdated=" & notUpdated)
		        elseif CDate(END_DATEarray(i)) <= CDate(EFF_DATEarray(i)) Then
			        ''go back to prior page with error message
			        Session("Err") = "**** ERROR: No updates/entries have been made. Make sure all records have valid dates for effective/end dates<br />and end dates are later than effective dates. ****"
			        Response.Redirect("PlayerRefCertification.asp?FILTER=" & vFILTER & "&notupdated=" & notUpdated)
		        end if
				




			End If

		Next

		'check to see if the same certification info entered already exists -
		''if not, add a new record (leave old one for historical purposes)'
		For i = 0 to UBound(IDarray)

			sql = 	"SELECT PERSON_ID, " & _
					"EFF_DATE, " & _
					"END_DATE, " & _
					"CERT_ID, " & _
					"NOTES " & _
					"FROM REFEREE_CERTIFICATION_TBL " & _
					"WHERE PERSON_ID = '" & IDarray(i) & "' " & _
					"AND CERT_ID = '" & CERT_IDarray(i) & "' " & _
					"AND EFF_DATE = '" & EFF_DATEarray(i) & "' " & _
					"AND END_DATE = '" & END_DATEarray(i) & "'"
			rw(sql)
			set rs = cn.Execute(sql)

			if rs.EOF then

				''dont insert if it has a missing eff date, certification

				if not (EFF_DATEarray(i) = "" or END_DATEarray(i) = "" or CERT_IDarray(i) = "-1") then

					sql = "INSERT INTO REFEREE_CERTIFICATION_TBL(PERSON_ID,  "_
						& "EFF_DATE, "_
						& "END_DATE, "_
						& "CERT_ID, "_
						& "NOTES) "_
						& "VALUES("_
						& "'" & IDarray(i) & "', "_
						& "'" & EFF_DATEarray(i) & "', "_
						& "'" & END_DATEarray(i) & "', "_
						& "'" & CERT_IDarray(i) & "', "_
						& "'" & NOTESarray(i) & "')"

					''rw(sql & "<br />")
					''Response.End

					cn.Execute(sql)

				end if

			''update notes if they arent blank and this record already exists
			elseif NOTESarray(i) <> "" AND CInt(rs("PERSON_ID")) = CInt(IDarray(i)) AND CStr(rs("EFF_DATE")) = CStr(EFF_DATEarray(i)) then

				sql = 	"UPDATE REFEREE_CERTIFICATION_TBL " & _
						"SET NOTES = '" & NOTESarray(i)  & "' " & _
						"WHERE PERSON_ID = '" & IDarray(i) & "' " & _
						"AND EFF_DATE = '" & EFF_DATEarray(i) & "'"
				cn.Execute(sql)

			end if

		Next

		call closeCNConnection()
		Session("admin") = "modify"

	end if

	Response.Redirect("PlayerRefCertification.asp?FILTER=" & vFILTER & "&notupdated=" & notUpdated)

'******************************************'

Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function
%>