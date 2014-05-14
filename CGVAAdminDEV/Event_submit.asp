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
	dim LOCATION_CD, LOCATION_CDarray
	dim EVENT_SHORT_DESC, EVENT_SHORT_DESCarray
	dim EVENT_LONG_DESC, EVENT_LONG_DESCarray
	dim EVENT_START_DATE, EVENT_START_DATEarray
	dim EVENT_END_DATE, EVENT_END_DATEarray
	dim EVENT_TYPE_ID, EVENT_TYPE_IDarray
	dim ACTIVE_EVENT_IND, ACTIVE_EVENT_INDarray
	dim WEEK_NUM_DISPLAY_IND, WEEK_NUM_DISPLAY_INDarray
	dim EVENT_IMG, EVENT_IMGarray

	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Event" then
		EVENT_CD		 		= Replace(Request("EVENT_CD"),"'","''")
		LOCATION_CD		 		= Replace(Request("LOCATION_CD"),"'","''")
		EVENT_SHORT_DESC 		= Replace(Request("EVENT_SHORT_DESC"),"'","''")
		EVENT_LONG_DESC 		= Replace(Request("EVENT_LONG_DESC"),"'","''")
		EVENT_START_DATE 		= Replace(Request("EVENT_START_DATE"),"'","''")
		EVENT_END_DATE 			= Replace(Request("EVENT_END_DATE"),"'","''")
		EVENT_TYPE_ID 			= Replace(Request("EVENT_TYPE_ID"),"'","''")
		WEEK_NUM_DISPLAY_IND	= Replace(Request("WEEK_NUM_DISPLAY_IND"),"'","''")
		ACTIVE_EVENT_IND 		= Replace(Request("ACTIVE_EVENT_IND"),"'","''")
		EVENT_IMG 				= Replace(Request("EVENT_IMG"),"'","''")

		'check to see if the event code entered exists'
		''not done yet
		''checkEvent()

		if exists <> "Y" then
			sql = "INSERT INTO EVENT_TBL("_
				& "EVENT_CD, "_
				& "LOCATION_CD, "_
				& "EVENT_SHORT_DESC, "_
				& "EVENT_LONG_DESC, "_
				& "EVENT_START_DATE, "_
				& "EVENT_END_DATE, "_
				& "EVENT_TYPE_ID, "_
				& "EVENT_IMG, "_
				& "ACTIVE_EVENT_IND, "_
				& "WEEK_NUM_DISPLAY_IND)"_
				& " VALUES("_
				& "'" & EVENT_CD & "', "_
				& "'" & LOCATION_CD & "', "_
				& "'" & EVENT_SHORT_DESC & "', "_
				& "'" & EVENT_LONG_DESC & "', "_
				& "'" & EVENT_START_DATE & "', "_
				& "'" & EVENT_END_DATE & "', "_
				& "'" & EVENT_TYPE_ID & "', "_
				& "'" & EVENT_IMG & "', "_
				& "'" & ACTIVE_EVENT_IND & "', "_
				& "'" & WEEK_NUM_DISPLAY_IND & "')"

			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)
		else
			call CloseCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Event.asp")
		end if

		call CloseCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Event.asp")


	ElseIf Request("submitChoice") = "Modify Event(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		LOCATION_CD = FixEmptyCell(Request("LOCATION_CD"))
		LOCATION_CDarray = split(LOCATION_CD,", ")

		EVENT_SHORT_DESC = FixEmptyCell(Request("EVENT_SHORT_DESC"))
		EVENT_SHORT_DESCarray = split(EVENT_SHORT_DESC,", ")

		EVENT_LONG_DESC = FixEmptyCell(Request("EVENT_LONG_DESC"))
		EVENT_LONG_DESCarray = split(EVENT_LONG_DESC,", ")

		EVENT_START_DATE = FixEmptyCell(Request("EVENT_START_DATE"))
		EVENT_START_DATEarray = split(EVENT_START_DATE,", ")

		EVENT_END_DATE = FixEmptyCell(Request("EVENT_END_DATE"))
		EVENT_END_DATEarray = split(EVENT_END_DATE,", ")

		EVENT_TYPE_ID = FixEmptyCell(Request("EVENT_TYPE_ID"))
		EVENT_TYPE_IDarray = split(EVENT_TYPE_ID,", ")

		WEEK_NUM_DISPLAY_IND = FixEmptyCell(Request("WEEK_NUM_DISPLAY_IND"))
		WEEK_NUM_DISPLAY_INDarray = split(WEEK_NUM_DISPLAY_IND,", ")

		ACTIVE_EVENT_IND = FixEmptyCell(Request("ACTIVE_EVENT_IND"))
		ACTIVE_EVENT_INDarray = split(ACTIVE_EVENT_IND,", ")

		EVENT_IMG = FixEmptyCell(Request("EVENT_IMG"))
		EVENT_IMGarray = split(EVENT_IMG,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			'check to see if the event code entered exists if they are updating the event code field'

			if IDarray(i) <> EVENT_CDarray(i) then
				checkEvent()
			else
				exists = "N"
			end if

			if exists <> "Y" then
				sql = "UPDATE EVENT_TBL SET "_
					& "EVENT_CD = '" & EVENT_CDarray(i) & "', "_
					& "LOCATION_CD = '" & LOCATION_CDarray(i) & "', "_
					& "EVENT_SHORT_DESC = '" & EVENT_SHORT_DESCarray(i) & "', "_
					& "EVENT_LONG_DESC = '" & EVENT_LONG_DESCarray(i) & "', "_
					& "EVENT_START_DATE = '" & EVENT_START_DATEarray(i) & "', "_
					& "EVENT_END_DATE = '" & EVENT_END_DATEarray(i) & "', "_
					& "EVENT_TYPE_ID = '" & EVENT_TYPE_IDarray(i) & "', "_
					& "EVENT_IMG = '" & EVENT_IMGarray(i) & "', "_
					& "ACTIVE_EVENT_IND = '" & ACTIVE_EVENT_INDarray(i) & "', "_
					& "WEEK_NUM_DISPLAY_IND = '" & WEEK_NUM_DISPLAY_INDarray(i) & "' "_
					& "WHERE EVENT_CD = '" & IDarray(i) & "'"

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
		Response.Redirect("Event.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Event.asp")
	End If


'******************************************'

Sub checkEvent()

	If Request("submitChoice") = "Modify Event(s)" then
		sql = "SELECT EVENT_CD FROM EVENT_TBL "_
			& "WHERE UPPER(EVENT_CD) = '" & UCase(EVENT_CDarray(i)) & "'"

	Else
		sql = "SELECT EVENT_CD FROM EVENT_TBL WHERE UPPER(EVENT_CD) = '" & UCase(EVENT_CD) & "' "
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