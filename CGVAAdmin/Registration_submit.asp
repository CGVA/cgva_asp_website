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
	''dim ID, IDarray
	dim EVENT_CD, EVENT_CDarray
	dim PERSON_ID, PERSON_IDarray
	dim DATE, DATEarray
	dim WAIVER_IND, WAIVER_INDarray
	dim OPEN_PLAY_IND, OPEN_PLAY_INDarray
	dim REGISTRATION_IND, REGISTRATION_INDarray
	dim DOLLARS_COLLECTED, DOLLARS_COLLECTEDarray
	dim DOLLARS_OFF_COUPON, DOLLARS_OFF_COUPONarray
	dim CHECK_AMT_COLLECTED, CHECK_AMT_COLLECTEDarray
	dim CHECK_NUM, CHECK_NUMarray
	dim NOTES, NOTESarray
	dim CAPTAIN_INTEREST_IND , CAPTAIN_INTEREST_INDarray
	dim VOLUNTEER_INTEREST_IND, VOLUNTEER_INTEREST_INDarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Registration" then
		EVENT_CD				= Replace(Request("EVENT_CD"),"'","''")
		PERSON_ID				= Replace(Request("PERSON_ID"),"'","''")
		vDATE					= Replace(Request("DATE"),"'","''")
		WAIVER_IND				= Replace(Request("WAIVER_IND"),"'","''")
		OPEN_PLAY_IND			= Replace(Request("OPEN_PLAY_IND"),"'","''")
		PLAY_UPDOWN_IND			= Replace(Request("PLAY_UPDOWN_IND"),"'","''")
		REGISTRATION_IND		= Replace(Request("REGISTRATION_IND"),"'","''")
		DOLLARS_COLLECTED		= Replace(Request("DOLLARS_COLLECTED"),"'","''")
		DOLLARS_OFF_COUPON		= Replace(Request("DOLLARS_OFF_COUPON"),"'","''")
		CHECK_AMT_COLLECTED		= Replace(Request("CHECK_AMT_COLLECTED"),"'","''")
		CHECK_NUM				= Replace(Request("CHECK_NUM"),"'","''")
		NOTES					= Replace(Request("NOTES"),"'","''")
		CAPTAIN_INTEREST_IND	= Replace(Request("CAPTAIN_INTEREST_IND"),"'","''")
		VOLUNTEER_INTEREST_IND	= Replace(Request("VOLUNTEER_INTEREST_IND"),"'","''")


		'check to see if the first/last name combo entered exists'
		checkRegistration()

		if exists <> "Y" then
			sql = "INSERT INTO REGISTRATION_TBL("_
				& "EVENT_CD, "_
				& "PERSON_ID, "_
				& "[DATE], "_
				& "WAIVER_IND, "_
				& "OPEN_PLAY_IND, "_
				& "PLAY_UPDOWN_IND, "_
				& "REGISTRATION_IND, "_
				& "DOLLARS_COLLECTED, "_
				& "DOLLARS_OFF_COUPON, "_
				& "CHECK_AMT_COLLECTED, "_
				& "CHECK_NUM, "_
				& "NOTES, "_
				& "CAPTAIN_INTEREST_IND, "_
				& "VOLUNTEER_INTEREST_IND)"_
				& " VALUES("_
				& "'" & EVENT_CD & "', "_
				& "'" & PERSON_ID & "', "_
				& "IsNull('" & vDATE & "',''), "_
				& "'" & WAIVER_IND & "', "_
				& "'" & OPEN_PLAY_IND & "', "_
				& "'" & PLAY_UPDOWN_IND & "', "_
				& "'" & REGISTRATION_IND & "', "_
				& "'" & DOLLARS_COLLECTED & "', "_
				& "'" & DOLLARS_OFF_COUPON & "', "_
				& "'" & CHECK_AMT_COLLECTED & "', "_
				& "'" & CHECK_NUM & "', "_
				& "'" & NOTES & "', "_
				& "'" & CAPTAIN_INTEREST_IND & "', "_
				& "'" & VOLUNTEER_INTEREST_IND & "')"


			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call closeRSCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Registration.asp")
		end if

		call closeRSCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Registration.asp")


	ElseIf Request("submitChoice") = "Modify Registration(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		PERSON_ID = FixEmptyCell(Request("PERSON_ID"))
		PERSON_IDarray = split(PERSON_ID,", ")

		vDATE = FixEmptyCell(Request("DATE"))
		vDATEarray = split(vDATE,", ")

		WAIVER_IND = FixEmptyCell(Request("WAIVER_IND"))
		WAIVER_INDarray = split(WAIVER_IND,", ")

		OPEN_PLAY_IND = FixEmptyCell(Request("OPEN_PLAY_IND"))
		OPEN_PLAY_INDarray = split(OPEN_PLAY_IND,", ")

		PLAY_UPDOWN_IND = FixEmptyCell(Request("PLAY_UPDOWN_IND"))
		PLAY_UPDOWN_INDarray = split(PLAY_UPDOWN_IND,", ")

		REGISTRATION_IND = FixEmptyCell(Request("REGISTRATION_IND"))
		REGISTRATION_INDarray = split(REGISTRATION_IND,", ")

		DOLLARS_COLLECTED = FixEmptyCell(Request("DOLLARS_COLLECTED"))
		DOLLARS_COLLECTEDarray = split(DOLLARS_COLLECTED,", ")

		DOLLARS_OFF_COUPON = FixEmptyCell(Request("DOLLARS_OFF_COUPON"))
		DOLLARS_OFF_COUPONarray = split(DOLLARS_OFF_COUPON,", ")

		CHECK_AMT_COLLECTED = FixEmptyCell(Request("CHECK_AMT_COLLECTED"))
		CHECK_AMT_COLLECTEDarray = split(CHECK_AMT_COLLECTED,", ")

		CHECK_NUM = FixEmptyCell(Request("CHECK_NUM"))
		CHECK_NUMarray = split(CHECK_NUM,", ")

		NOTES = FixEmptyCell(Request("NOTES"))
		NOTESarray = split(NOTES,", ")

		CAPTAIN_INTEREST_IND = FixEmptyCell(Request("CAPTAIN_INTEREST_IND"))
		CAPTAIN_INTEREST_INDarray = split(CAPTAIN_INTEREST_IND,", ")

		VOLUNTEER_INTEREST_IND = FixEmptyCell(Request("VOLUNTEER_INTEREST_IND"))
		VOLUNTEER_INTEREST_INDarray = split(VOLUNTEER_INTEREST_IND,", ")



		'check to see if duplicate exists'
		For i = 0 to UBound(PERSON_IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			checkRegistration()

			if exists <> "Y" then
				sql = "UPDATE REGISTRATION_TBL SET "_
					& "EVENT_CD = '" & EVENT_CDarray(i) & "', "_
					& "PERSON_ID = '" & PERSON_IDarray(i) & "', "_
					& "DATE = '" & vDATEarray(i) & "', "_
					& "WAIVER_IND = '" & WAIVER_INDarray(i) & "', "_
					& "OPEN_PLAY_IND = '" & OPEN_PLAY_INDarray(i) & "', "_
					& "PLAY_UPDOWN_IND = '" & PLAY_UPDOWN_INDarray(i) & "', "_
					& "REGISTRATION_IND = '" & REGISTRATION_INDarray(i) & "', "_
					& "DOLLARS_COLLECTED = '" & DOLLARS_COLLECTEDarray(i) & "', "_
					& "DOLLARS_OFF_COUPON = '" & DOLLARS_OFF_COUPONarray(i) & "', "_
					& "CHECK_AMT_COLLECTED = '" & CHECK_AMT_COLLECTEDarray(i) & "', "_
					& "CHECK_NUM = '" & CHECK_NUMarray(i) & "', "_
					& "NOTES = '" & Replace(NOTESarray(i),"'","''") & "', "_
					& "CAPTAIN_INTEREST_IND = '" & CAPTAIN_INTEREST_INDarray(i) & "', "_
					& "VOLUNTEER_INTEREST_IND = '" & VOLUNTEER_INTEREST_INDarray(i) & "' "_
					& "WHERE ID = '" & IDarray(i) & "'"

				'rw(sql & "<br />")
				'Response.End

				cn.Execute(sql)
			else
				notUpdated = notUpdated & " Person ID: " & PERSON_IDarray(i) & "<br />"
			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Registration.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Registration.asp")
	End If


'******************************************'

Sub checkRegistration()

	If Request("submitChoice") = "Modify Registration(s)" then
		sql = "SELECT PERSON_ID FROM REGISTRATION_TBL "_
			& "WHERE PERSON_ID = '" & UCase(PERSON_IDarray(i)) & "' "_
			& "AND EVENT_CD = '" & UCase(EVENT_CDarray(i)) & "' "_
			& "AND [DATE] = '" & UCase(vDATEarray(i)) & "' "_
			& "AND [OPEN_PLAY_IND] = '" & UCase(OPEN_PLAY_INDarray(i)) & "' "_
			& "AND [REGISTRATION_IND] = '" & UCase(REGISTRATION_INDarray(i)) & "' "_
			& "AND ID <> '" & IDarray(i) & "'"

	Else

		''1/13/2008 - modify this so that a person also cant be registered twice for the same event
		sql = "SELECT PERSON_ID FROM REGISTRATION_TBL "_
			& "WHERE PERSON_ID = '" & PERSON_ID & "' "_
			& "AND EVENT_CD = '" & EVENT_CD & "' "_
			& "AND [DATE] = '" & vDATE & "' "_
			& "AND [OPEN_PLAY_IND] = '" & OPEN_PLAY_IND & "' "_
			& "AND [REGISTRATION_IND] = '" & REGISTRATION_IND & "' "_
			& "UNION "_
			& "SELECT PERSON_ID FROM REGISTRATION_TBL "_
			& "WHERE PERSON_ID = '" & PERSON_ID & "' "_
			& "AND EVENT_CD = '" & EVENT_CD & "' "_
			& "AND [REGISTRATION_IND] = 'Y'"
	End If

	rw(sql & "<br />")
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