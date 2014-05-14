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
	dim RATING_PERSON_ID1, RATING_PERSON_ID1array
	dim RATING_PERSON_ID2, RATING_PERSON_ID2array
	dim RATING_PERSON_ID3, RATING_PERSON_ID3array
	dim RATING_SCORE, RATING_SCOREarray
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

	RATING_PERSON_ID1 		= FixEmptyCell(Replace(Request("RATING_PERSON_ID1"),"'","''"))
	RATING_PERSON_ID1array	= split(RATING_PERSON_ID1,", ")

	RATING_PERSON_ID2		= FixEmptyCell(Replace(Request("RATING_PERSON_ID2"),"'","''"))
	RATING_PERSON_ID2array	= split(RATING_PERSON_ID2,", ")

	RATING_PERSON_ID3		= FixEmptyCell(Replace(Request("RATING_PERSON_ID3"),"'","''"))
	RATING_PERSON_ID3array	= split(RATING_PERSON_ID3,", ")

	RATING_SCORE   			= FixEmptyCell(Replace(Request("RATING_SCORE"),"'","''"))
	RATING_SCOREarray		= split(RATING_SCORE,", ")

	NOTES					= FixEmptyCell(Replace(Request("NOTES"),"'","''"))
	NOTESarray				= split(NOTES,", ")



	'check to see if the same ratings info entered already exists -
	''if not, add a new record (leave old one for historical purposes)'
	For i = 0 to UBound(IDarray)

		sql = 	"SELECT PERSON_ID, " & _
				"EFF_DATE, " & _
				"RATING_PERSON_ID1, " & _
				"RATING_PERSON_ID2, " & _
				"RATING_PERSON_ID3, " & _
				"RATING_SCORE, " & _
				"NOTES " & _
				"FROM RATINGS_TBL " & _
				"WHERE PERSON_ID = '" & IDarray(i) & "' " & _
				"AND EFF_DATE = '" & EFF_DATEarray(i) & "' " & _
				"AND RATING_PERSON_ID1 = '" & RATING_PERSON_ID1array(i) & "'"
''				"AND RATING_SCORE = '" & RATING_SCOREarray(i) & "'"
		rw(sql)
		set rs = cn.Execute(sql)

		if rs.EOF then
			''rw(RATING_PERSON_ID3array(i) & "<br />")
			''rw(CStr(RATING_PERSON_ID3array(i)) & "<br />")
			''rw(Replace(CStr(RATING_PERSON_ID3array(i)),"-1","NULL") & "<br />")
			''Response.End

			''RATING_PERSON_ID1array(i) = Replace(CStr(RATING_PERSON_ID1array(i)),"-1","NULL")
			RATING_PERSON_ID2array(i) = Replace(CStr(RATING_PERSON_ID2array(i)),"-1","NULL")
			RATING_PERSON_ID3array(i) = Replace(CStr(RATING_PERSON_ID3array(i)),"-1","NULL")

			''dont insert if it has a missing eff date, ratings person 1 or rating score
			''rw("D:" & EFF_DATEarray(i) & "<br />")
			''rw("RP:" & RATING_PERSON_ID1array(i) & "<br />")
			''rw("R:" & RATING_SCOREarray(i) & "<br />")
			''Response.End
			if not (EFF_DATEarray(i) = "" or RATING_PERSON_ID1array(i) = "-1" or RATING_SCOREarray(i) = "") then

				sql = "INSERT INTO RATINGS_TBL(PERSON_ID,  "_
					& "EFF_DATE, "_
					& "RATING_PERSON_ID1, "_
					& "RATING_PERSON_ID2, "_
					& "RATING_PERSON_ID3, "_
					& "RATING_SCORE, "_
					& "NOTES) "_
					& "VALUES("_
					& "'" & IDarray(i) & "', "_
					& "'" & EFF_DATEarray(i) & "', "_
					& "" & RATING_PERSON_ID1array(i) & ", "_
					& "" & RATING_PERSON_ID2array(i) & ", "_
					& "" & RATING_PERSON_ID3array(i) & ", "_
					& "'" & RATING_SCOREarray(i) & "', "_
					& "'" & NOTESarray(i) & "')"

				''rw(sql & "<br />")
				''Response.End

				cn.Execute(sql)

			end if

		''update notes if they arent blank and this record already exists
		elseif NOTESarray(i) <> "" AND CInt(rs("PERSON_ID")) = CInt(IDarray(i)) AND CInt(rs("RATING_PERSON_ID1")) = CInt(RATING_PERSON_ID1array(i)) AND CStr(rs("EFF_DATE")) = CStr(EFF_DATEarray(i)) then
		''AND CDbl(rs("RATING_SCORE")) = CDbl(RATING_SCOREarray(i)) then
			sql = 	"UPDATE RATINGS_TBL " & _
					"SET NOTES = '" & NOTESarray(i)  & "' " & _
					"WHERE PERSON_ID = '" & IDarray(i) & "' " & _
					"AND EFF_DATE = '" & EFF_DATEarray(i) & "' " & _
					"AND [RATING_PERSON_ID1] = '" & RATING_PERSON_ID1array(i) & "'"
					''"AND RATING_SCORE = '" & RATING_SCOREarray(i) & "'"
			cn.Execute(sql)

			'rw(rs("PERSON_ID") & ", " & IDarray(i) & "<br />")
			'rw(rs("RATING_PERSON_ID1") & ", " & RATING_PERSON_ID1array(i) & "<br />")
			'rw(rs("EFF_DATE") & ", " & EFF_DATEarray(i) & "<br />")
			'rw(rs("RATING_SCORE") & ", " & RATING_SCOREarray(i) & "<br />")
			'Response.End
		end if


	Next

''RATING_PERSON_ID1array(i) = Replace(CStr(RATING_PERSON_ID1array(i)),"-1","NULL")
''RATING_PERSON_ID2array(i) = Replace(CStr(RATING_PERSON_ID2array(i)),"-1","NULL")
''RATING_PERSON_ID3array(i) = Replace(CStr(RATING_PERSON_ID3array(i)),"-1","NULL")

''sql = "UPDATE RATINGS_TBL SET "_
''	& "EFF_DATE 			= '" & EFF_DATEarray(i) & "', "_
''	& "RATING_PERSON_ID1 	= " & RATING_PERSON_ID1array(i) & ", "_
''	& "RATING_PERSON_ID2 	= " & RATING_PERSON_ID2array(i) & ", "_
''	& "RATING_PERSON_ID3 	= " & RATING_PERSON_ID3array(i) & ", "_
''	& "RATING_SCORE 		= '" & RATING_SCOREarray(i) & "', "_
''	& "NOTES 				= '" & NOTESarray(i) & "' "_
''	& "WHERE PERSON_ID 		= '" & IDarray(i) & "'"
''		else
''		Response.End

	call closeCNConnection()
	Session("admin") = "modify"
	Response.Redirect("PlayerRatings.asp?FILTER=" & vFILTER & "&notupdated=" & notUpdated)

'******************************************'

Function FixEmptyCell(value)

	If value = "" Then
		value = " "
	End If

	FixEmptyCell = value
End Function
%>