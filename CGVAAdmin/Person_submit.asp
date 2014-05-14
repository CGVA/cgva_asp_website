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
	dim PERSON_ID, PERSON_IDarray
	dim FIRST_NAME, FIRST_NAMEarray
	dim LAST_NAME, LAST_NAMEarray
	dim EMAIL, EMAILarray
	dim BIRTH_DATE, BIRTH_DATEarray
	dim GENDER, GENDERarray
	dim PRIMARY_PHONE_NUM, PRIMARY_PHONE_NUMarray
	dim PHONE2, PHONE2array
	dim PHONE3,PHONE3array
	dim ADDRESS_LINE1, ADDRESS_LINE1array
	dim ADDRESS_LINE2, ADDRESS_LINE2array
	dim CITY, CITYarray
	dim STATE, STATEarray
	dim ZIP, ZIParray
	dim EMERGENCY_FIRST_NAME , EMERGENCY_FIRST_NAMEarray
	dim EMERGENCY_LAST_NAME, EMERGENCY_LAST_NAMEarray
	dim EMERGENCY_PHONE, EMERGENCY_PHONEarray
	dim SUPPRESS_EMAIL_IND, SUPPRESS_EMAIL_INDarray
	dim SUPPRESS_SNAIL_MAIL_IND, SUPPRESS_SNAIL_MAIL_INDarray
	dim SUPPRESS_LAST_NAME_IND, SUPPRESS_LAST_NAME_INDarray
	dim FIRST_CONTACT_ID, FIRST_CONTACT_IDarray
	dim PHOTO_FILENAME, PHOTO_FILENAMEarray
	dim DATE_ADDED, DATE_ADDEDarray
	dim USER_ADDED, USER_ADDEDarray
''	dim DATE_UPDATED, array
''	dim USER_UPDATED, array
''	dim LOGICAL_DELETE_IND, array
''	dim DATE_DELETED, array
''	dim USER_DELETED, array
	dim NAGVA_RATING, NAGVA_RATINGarray
	dim LASTDIG_IND, LASTDIG_INDarray


	If Session("ACCESS") = "" then
		Session("Err") = "Your session has timed out. Please log in again."
		Response.Redirect("Adminindex.asp")
	ElseIf Not Instr(Session("ACCESS"),"ADMIN") > 0 Then
		Session("Err") = "You do not have access to view the requested page."
		Response.Redirect("Adminindex.asp")
	End If

	If Request("submitChoice") = "Add Person" then
		FIRST_NAME				= Replace(Request("FIRST_NAME"),"'","''")
		LAST_NAME				= Replace(Request("LAST_NAME"),"'","''")
		EMAIL					= Replace(Request("EMAIL"),"'","''")
		BIRTH_DATE				= Replace(Request("BIRTH_DATE"),"'","''")
		GENDER					= Replace(Request("GENDER"),"'","''")
		PRIMARY_PHONE_NUM		= Replace(Request("PRIMARY_PHONE_NUM"),"'","''")
		PHONE2					= Replace(Request("PHONE2"),"'","''")
		PHONE3					= Replace(Request("PHONE3"),"'","''")
		ADDRESS_LINE1			= Replace(Request("ADDRESS_LINE1"),"'","''")
		ADDRESS_LINE2			= Replace(Request("ADDRESS_LINE2"),"'","''")
		CITY					= Replace(Request("CITY"),"'","''")
		STATE					= Replace(Request("STATE"),"'","''")
		ZIP						= Replace(Request("ZIP"),"'","''")
		EMERGENCY_FIRST_NAME	= Replace(Request("EMERGENCY_FIRST_NAME"),"'","''")
		EMERGENCY_LAST_NAME		= Replace(Request("EMERGENCY_LAST_NAME"),"'","''")
		EMERGENCY_PHONE			= Replace(Request("EMERGENCY_PHONE"),"'","''")
		SUPPRESS_EMAIL_IND		= Replace(Request("SUPPRESS_EMAIL_IND"),"'","''")
		SUPPRESS_SNAIL_MAIL_IND	= Replace(Request("SUPPRESS_SNAIL_MAIL_IND"),"'","''")
		SUPPRESS_LAST_NAME_IND	= Replace(Request("SUPPRESS_LAST_NAME_IND"),"'","''")
		FIRST_CONTACT_ID		= Replace(Request("FIRST_CONTACT_ID"),"'","''")
		PHOTO_FILENAME			= Replace(Request("PHOTO_FILENAME"),"'","''")
''		DATE_ADDED				= Date
		USER_ADDED				= "1"
		''DATE_UPDATED			= Replace(Request(""),"'","''")
		''USER_UPDATED			= Replace(Request(""),"'","''")
		''LOGICAL_DELETE_IND	= Replace(Request(""),"'","''")
		''DATE_DELETED			= Replace(Request(""),"'","''")
		''USER_DELETED			= Replace(Request(""),"'","''")
		NAGVA_RATING			= Replace(Request("NAGVA_RATING"),"'","''")
		LASTDIG_IND				= Replace(Request("LASTDIG_IND"),"'","''")


		'check to see if the first/last name combo entered exists'
		''JPC 12/7/2007 - two people with same name are OK, they will just have diff IDs
		''JPC 8/6/2008 dont want two people with same name, do something different
		checkPerson()

		if exists <> "Y" then
			sql = "INSERT INTO db_accessadmin.PERSON_TBL("_
				& "FIRST_NAME, "_
				& "LAST_NAME, "_
				& "EMAIL, "_
				& "BIRTH_DATE, "_
				& "GENDER, "_
				& "PRIMARY_PHONE_NUM, "_
				& "[2ND_PHONE_NUM], "_
				& "[3RD_PHONE_NUM], "_
				& "ADDRESS_LINE1, "_
				& "ADDRESS_LINE2, "_
				& "CITY, "_
				& "STATE, "_
				& "ZIP, "_
				& "EMERGENCY_FIRST_NAME , "_
				& "EMERGENCY_LAST_NAME, "_
				& "EMERGENCY_PHONE, "_
				& "SUPPRESS_EMAIL_IND, "_
				& "SUPPRESS_SNAIL_MAIL_IND, "_
				& "SUPPRESS_LAST_NAME_IND, "_
				& "FIRST_CONTACT_ID, "_
				& "PHOTO_FILENAME, "_
				& "DATE_ADDED, "_
				& "USER_ADDED, "_
				& "NAGVA_RATING, "_
				& "LASTDIG_IND)"_
				& " VALUES("_
				& "'" & FIRST_NAME & "', "_
				& "'" & LAST_NAME & "', "_
				& "'" & EMAIL & "', "_
				& "IsNull('" & BIRTH_DATE & "',''), "_
				& "'" & GENDER & "', "_
				& "'" & PRIMARY_PHONE_NUM & "', "_
				& "'" & PHONE2 & "', "_
				& "'" & PHONE3 & "', "_
				& "'" & ADDRESS_LINE1 & "', "_
				& "'" & ADDRESS_LINE2 & "', "_
				& "'" & CITY & "', "_
				& "'" & STATE & "', "_
				& "'" & ZIP & "', "_
				& "'" & EMERGENCY_FIRST_NAME & "', "_
				& "'" & EMERGENCY_LAST_NAME & "', "_
				& "'" & EMERGENCY_PHONE & "', "_
				& "'" & SUPPRESS_EMAIL_IND & "', "_
				& "'" & SUPPRESS_SNAIL_MAIL_IND & "', "_
				& "'" & SUPPRESS_LAST_NAME_IND & "', "_
				& "'" & FIRST_CONTACT_ID & "', "_
				& "'" & PHOTO_FILENAME & "', "_
				& "getdate(), "_
				& "'" & USER_ADDED & "', "_
				& "'" & NAGVA_RATING & "', "_
				& "'" & LASTDIG_IND & "')"


			rw(sql & "<br />")
			''Response.End

			cn.Execute(sql)

			'get this user ID and add a record to the login table
			SQL = "select top 1 PERSON_ID "_
				& "FROM db_accessadmin.PERSON_TBL "_
				& "ORDER BY DATE_ADDED DESC"

			set rs = cn.Execute(SQL)
			if not rs.EOF then
				PERSON_ID = rs(0)
				uName = LCase(Left(FIRST_NAME,1) & LAST_NAME)
				uPassword = uName + CStr(DatePart("n",Now)) + CStr(DatePart("s",Now))
				SQL = "INSERT INTO USER_LOGIN_TBL(PERSON_ID,USERNAME,TEMP_PASSWORD) "_
					& "SELECT '" & PERSON_ID & "', '" & uName & "','" & uPassword & "'"

				rw(SQL & "<br />")
				''Response.End
				cn.Execute(SQL)

				'send MyCGVA email
				dim myMail
				Set myMail=CreateObject("CDO.Message")

				myMail.HTMLBody="What should the body read?<p />Registration for " & Session("clinicName") & " is complete."

                myMail.From = "cgva@cgva.org"
                myMail.To = EMAIL
                myMail.Bcc="support@cgva.org"
                myMail.Subject = "Welcome to MyCGVA!"
                myMail.HTMLBody			= "<div>Dear " & FIRST_NAME & ", " _
                                        & "<br /><br />" _
                                        & "The CGVA website team has rolled out a brand new online experience called MyCGVA! Initially, MyCGVA will provide a secure area on the CGVA website where you can:" _
                                        & "<br /><br />" _
                                        & "-Register/pay for CGVA leagues and tournaments" _
                                        & "<br />" _
                                        & "-Set up preferences" _
                                        & "<br />" _
                                        & "-Manage your contact information" _
                                        & "<br />" _
                                        & "-Reset your password and your security questions" _
                                        & "<br /><br />" _
                                        & "We have created an ID and a temporary password for you, which you must change the first time you logon to MyCGVA. " _
                                        & "<br /><br />" _
                                        & "User Name: " & uName _
                                        & "<br />" _
                                        & "Temporary Password: " & uPassword _
                                        & "<br /><br />" _
                                        & "Please use the link below to go to the login page, change your password and begin using MyCGVA! " _
                                        & "<br /><br />" _
                                        & "<a href='http://www.cgva.org/MyCGVA/Login.aspx'>http://www.cgva.org/MyCGVA/Login.aspx</a>" _
                                        & "<br /><br />" _
                                        & "Note that information that we have on file for you may be out-dated or incorrect.  Feel free to go into the 'My Profile' area of the site to change your information." _
                                       & "<br /><br />" _
                                        & "Please add cgva@cgva.org and support@cgva.org to your address book so you receive all future e-blasts about CGVA and any messages related to your account. " _
                                        & "<br /><br />" _
                                        & "If you are not interested in having an ID on the CGVA website, please send an email to support@cgva.org and we’ll deactivate your account." _
                                        & "<br /><br />" _
                                        & "Sincerely," _
                                        & "<br /><br />" _
                                        & "The CGVA Website Team" _
                                        & "<br /><br />" _
                                        & "For General info/questions about CGVA send emails to cgva@cgva.org " _
                                        & "<br />" _
                                        & "For Support issues/problems with the CGVA Website send emails to support@cgva.org </div>"


				myMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

				'Name or IP of remote SMTP server
				myMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserver") _
				="relay-hosting.secureserver.net"

				'Server port
				myMail.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserverport") _
				=25

				myMail.Configuration.Fields.Update
				myMail.Send
				set myMail=nothing

			end if

		else
			call closeCNConnection()
			Session("admin") = "insertFail"
			Response.Redirect("Person.asp")
		end if

		call closeCNConnection()
		Session("admin") = "insert"
		Response.Redirect("Person.asp")


	ElseIf Request("submitChoice") = "Modify Person(s)" then

		''rw("here")
		''Response.End

		'determine which record to modify based on the ID number of the record'
		ID = Request("ID")
		IDarray = split(ID,", ")

		EVENT_CD = FixEmptyCell(Request("EVENT_CD"))
		EVENT_CDarray = split(EVENT_CD,", ")

		FIRST_NAME					= FixEmptyCell(Replace(Request("FIRST_NAME"),"'","''"))
		FIRST_NAMEarray				= split(FIRST_NAME,", ")

		LAST_NAME					= FixEmptyCell(Replace(Request("LAST_NAME"),"'","''"))
		LAST_NAMEarray				= split(LAST_NAME,", ")

		EMAIL						= FixEmptyCell(Replace(Request("EMAIL"),"'","''"))
		EMAILarray					= split(EMAIL,", ")

		BIRTH_DATE					= FixEmptyCell(Replace(Request("BIRTH_DATE"),"'","''"))
		BIRTH_DATEarray				= split(BIRTH_DATE,", ")

		GENDER						= FixEmptyCell(Replace(Request("GENDER"),"'","''"))
		GENDERarray					= split(GENDER,", ")

		PRIMARY_PHONE_NUM			= FixEmptyCell(Replace(Request("PRIMARY_PHONE_NUM"),"'","''"))
		PRIMARY_PHONE_NUMarray		= split(PRIMARY_PHONE_NUM,", ")

		PHONE2						= FixEmptyCell(Replace(Request("PHONE2"),"'","''"))
		PHONE2array					= split(PHONE2,", ")

		PHONE3						= FixEmptyCell(Replace(Request("PHONE3"),"'","''"))
		PHONE3array					= split(PHONE3,", ")

		ADDRESS_LINE1				= FixEmptyCell(Replace(Request("ADDRESS_LINE1"),"'","''"))
		ADDRESS_LINE1array			= split(ADDRESS_LINE1,", ")

		ADDRESS_LINE2				= FixEmptyCell(Replace(Request("ADDRESS_LINE2"),"'","''"))
		ADDRESS_LINE2array			= split(ADDRESS_LINE2,", ")

		CITY						= FixEmptyCell(Replace(Request("CITY"),"'","''"))
		CITYarray					= split(CITY,", ")

		STATE						= FixEmptyCell(Replace(Request("STATE"),"'","''"))
		STATEarray					= split(STATE,", ")

		ZIP							= FixEmptyCell(Replace(Request("ZIP"),"'","''"))
		ZIParray					= split(ZIP,", ")

		EMERGENCY_FIRST_NAME		= FixEmptyCell(Replace(Request("EMERGENCY_FIRST_NAME"),"'","''"))
		EMERGENCY_FIRST_NAMEarray	= split(EMERGENCY_FIRST_NAME,", ")

		EMERGENCY_LAST_NAME			= FixEmptyCell(Replace(Request("EMERGENCY_LAST_NAME"),"'","''"))
		EMERGENCY_LAST_NAMEarray	= split(EMERGENCY_LAST_NAME,", ")

		EMERGENCY_PHONE				= FixEmptyCell(Replace(Request("EMERGENCY_PHONE"),"'","''"))
		EMERGENCY_PHONEarray		= split(EMERGENCY_PHONE,", ")

		SUPPRESS_EMAIL_IND			= FixEmptyCell(Replace(Request("SUPPRESS_EMAIL_IND"),"'","''"))
		SUPPRESS_EMAIL_INDarray		= split(SUPPRESS_EMAIL_IND,", ")

		SUPPRESS_SNAIL_MAIL_IND		= FixEmptyCell(Replace(Request("SUPPRESS_SNAIL_MAIL_IND"),"'","''"))
		SUPPRESS_SNAIL_MAIL_INDarray= split(SUPPRESS_SNAIL_MAIL_IND,", ")

		SUPPRESS_LAST_NAME_IND		= FixEmptyCell(Replace(Request("SUPPRESS_LAST_NAME_IND"),"'","''"))
		SUPPRESS_LAST_NAME_INDarray	= split(SUPPRESS_LAST_NAME_IND,", ")

		FIRST_CONTACT_ID			= FixEmptyCell(Replace(Request("FIRST_CONTACT_ID"),"'","''"))
		FIRST_CONTACT_IDarray		= split(FIRST_CONTACT_ID,", ")

		PHOTO_FILENAME				= FixEmptyCell(Replace(Request("PHOTO_FILENAME"),"'","''"))
		PHOTO_FILENAMEarray			= split(PHOTO_FILENAME,", ")

''		DATE_ADDED					= FixEmptyCell(Replace(Request("DATE_ADDED"),"'","''"))
''		DATE_ADDEDarray				= split(DATE_ADDED,", ")

''		USER_ADDED					= FixEmptyCell(Replace(Request("USER_ADDED"),"'","''"))
''		USER_ADDEDarray				= split(USER_ADDED,", ")

''		DATE_UPDATED				= FixEmptyCell(Replace(Request("DATE_UPDATED"),"'","''"))
''		DATE_UPDATEDarray			= split(DATE_UPDATED,", ")

''		USER_UPDATED				= FixEmptyCell(Replace(Request("USER_UPDATED"),"'","''"))
''		USER_UPDATEDarray			= split(USER_UPDATED,", ")

		LASTDIG_IND					= FixEmptyCell(Replace(Request("LASTDIG_IND"),"'","''"))
		LASTDIG_INDarray			= split(LASTDIG_IND,", ")

		LOGICAL_DELETE_IND			= FixEmptyCell(Replace(Request("LOGICAL_DELETE_IND"),"'","''"))
		LOGICAL_DELETE_INDarray		= split(LOGICAL_DELETE_IND,", ")

''		DATE_DELETED				= FixEmptyCell(Replace(Request("DATE_DELETED"),"'","''"))
''		DATE_DELETEDarray			= split(DATE_DELETED,", ")

''		USER_DELETED				= FixEmptyCell(Replace(Request("USER_DELETED"),"'","''"))
''		USER_DELETEDarray			= split(USER_DELETED,", ")

		NAGVA_RATING				= FixEmptyCell(Replace(Request("NAGVA_RATING"),"'","''"))
		NAGVA_RATINGarray			= split(NAGVA_RATING,", ")


		'check to see if the login ID entered exists'
		For i = 0 to UBound(IDarray)
''			rw(i & ": " & LOCATION_CDarray(i) & " ")
			''checkPerson()

''			if exists <> "Y" then

				''verify this user isnt on an active team if delete indicator is 'Y'
				if LOGICAL_DELETE_INDarray(i) = "Y" then
					sql = "SELECT a.PERSON_ID "_
						& "FROM REGISTRATION_TBL a "_
						& "LEFT JOIN EVENT_TBL b ON a.EVENT_CD = b.EVENT_CD "_
						& "WHERE b.ACTIVE_EVENT_IND = 'Y' "_
						& "AND a.REGISTRATION_IND = 'Y' "_
						& "AND a.PERSON_ID = '" & IDarray(i) & "'"
					set rs = cn.Execute(sql)

					if not rs.EOF then
						notUpdated = notUpdated & " " & IDarray(i) & "(member of an active team)<br />"
						update = "N"
					else
						update = "Y"
					end if

				end if

				''process with update if 2 conditions met
				if LOGICAL_DELETE_INDarray(i) = "N" or update = "Y" then
					sql = "UPDATE db_accessadmin.PERSON_TBL SET "_
						& "FIRST_NAME = '" & FIRST_NAMEarray(i) & "', "_
						& "LAST_NAME = '" & LAST_NAMEarray(i) & "', "_
						& "EMAIL = '" & EMAILarray(i) & "', "_
						& "BIRTH_DATE = '" & BIRTH_DATEarray(i) & "', "_
						& "GENDER = '" & GENDERarray(i) & "', "_
						& "PRIMARY_PHONE_NUM = '" & PRIMARY_PHONE_NUMarray(i) & "', "_
						& "[2ND_PHONE_NUM] = '" & PHONE2array(i) & "', "_
						& "[3RD_PHONE_NUM] = '" & PHONE3array(i) & "', "_
						& "ADDRESS_LINE1 = '" & ADDRESS_LINE1array(i) & "', "_
						& "ADDRESS_LINE2 = '" & ADDRESS_LINE2array(i) & "', "_
						& "CITY = '" & CITYarray(i) & "', "_
						& "STATE = '" & STATEarray(i) & "', "_
						& "ZIP = '" & ZIParray(i) & "', "_
						& "EMERGENCY_FIRST_NAME = '" & EMERGENCY_FIRST_NAMEarray(i) & "', "_
						& "EMERGENCY_LAST_NAME = '" & EMERGENCY_LAST_NAMEarray(i) & "', "_
						& "EMERGENCY_PHONE = '" & EMERGENCY_PHONEarray(i) & "', "_
						& "SUPPRESS_EMAIL_IND = '" & SUPPRESS_EMAIL_INDarray(i) & "', "_
						& "SUPPRESS_SNAIL_MAIL_IND = '" & SUPPRESS_SNAIL_MAIL_INDarray(i) & "', "_
						& "SUPPRESS_LAST_NAME_IND = '" & SUPPRESS_LAST_NAME_INDarray(i) & "', "_
						& "FIRST_CONTACT_ID = '" & FIRST_CONTACT_IDarray(i) & "', "_
						& "PHOTO_FILENAME = '" & PHOTO_FILENAMEarray(i) & "', "_
						& "USER_ADDED = '1', "_
						& "NAGVA_RATING = '" & NAGVA_RATINGarray(i) & "', "_
						& "LASTDIG_IND = '" & LASTDIG_INDarray(i) & "', "_
						& "LOGICAL_DELETE_IND = '" & LOGICAL_DELETE_INDarray(i) & "' "_
						& "WHERE PERSON_ID = '" & IDarray(i) & "'"

					rw(sql & "<br />")
					''Response.End

					cn.Execute(sql)

				end if

''			else
''				notUpdated = notUpdated & " " & IDarray(i)
''			end if

		Next

''		Response.End

		call closeCNConnection()
		Session("admin") = "modify"
		Response.Redirect("Person.asp?notupdated=" & notUpdated)

	Else
		rw("ERROR:" & Request("submitChoice"))
		Response.End
		Session("Err") = "<br /><br /><div align='center'><font class='cfontError10'><b>An error occurred. The record(s) did not get added/updated/deleted. Please try again.</b></font></div>"
		Response.Redirect("Person.asp")
	End If


'******************************************'

Sub checkPerson()

	If Request("submitChoice") = "Modify Person(s)" then
		sql = "SELECT PERSON_ID FROM db_accessadmin.PERSON_TBL "_
			& "WHERE LAST_NAME = '" & UCase(LAST_NAMEarray(i)) & "' "_
			& "AND FIRST_NAME = '" & UCase(FIRST_NAMEarray(i)) & "' "_
			& "AND PERSON_ID <> '" & IDarray(i) & "'"

	Else
		sql = "SELECT PERSON_ID FROM db_accessadmin.PERSON_TBL "_
			& "WHERE LAST_NAME = '" & UCase(LAST_NAME) & "' "_
			& "AND FIRST_NAME = '" & UCase(FIRST_NAME) & "'"
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