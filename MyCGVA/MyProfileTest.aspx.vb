Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System
Imports System.IO

Partial Class MyProfileTest
    Inherits System.Web.UI.Page

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Public personID As String = "0"

    Private userName, userName1, userName2, userName3 As String
    Private FNAME As String
    Private LNAME As String
    Private strEmail As String
    Private firstContactID As Integer = 0
    Private strGender As String
    Private strTshirt As String
    Private strNAGVA As String
    Private strPhone1 As String
    Private strPhone2 As String
    Private strPhone3 As String
    Private strBirthDate As String
    Private strADDRESS_LINE1 As String
    Private strADDRESS_LINE2 As String
    Private strCITY As String
    Private strSTATE As String
    Private strZIP As String
    Private strEMERGENCY_FIRST_NAME As String
    Private strEMERGENCY_LAST_NAME As String
    Private strEMERGENCY_PHONE As String
    Private strSUPPRESS_EMAIL_IND As String
    Private strSUPPRESS_SNAIL_MAIL_IND As String
    Private strSUPPRESS_LAST_NAME_IND As String


    'Private strPassword1 As String
    'Private strPassword2 As String
    'Private securityQuestion As Integer
    'Private securityAnswer As String

    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        'Dim myString As String

        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString
        Else
            personID = "0"
        End If

        If personID = "0" Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If

        ''Response.Write(myString)

        If Not IsPostBack Then

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()
            'get information for this PERSON_ID
            SQL = "SELECT *,IsNull([2ND_PHONE_NUM],'') as 'PHONE2', " _
                & "IsNull([3RD_PHONE_NUM],'') as 'PHONE3', " _
                & "IsNull(ADDRESS_LINE1,'') as 'ADD_LINE1', " _
                & "IsNull(ADDRESS_LINE2,'') as 'ADD_LINE2', " _
                & "IsNull(CITY,'') as 'vCITY', " _
                & "IsNull(ZIP,'') as 'vZIP', " _
                & "IsNull(STATE,'') as 'vSTATE', " _
                & "IsNull(EMERGENCY_FIRST_NAME,'') as 'vEMER_FNAME', " _
                & "IsNull(EMERGENCY_LAST_NAME,'') as 'vEMER_LNAME', " _
                & "IsNull(EMERGENCY_PHONE,'') as 'vEMER_PHONE', " _
                & "IsNull(NAGVA_RATING,'') as 'vNAGVA', " _
                & "IsNull(TSHIRT_SIZE,'') as 'vTSHIRT' " _
                & "FROM db_accessadmin.[PERSON_TBL] " _
                & "WHERE PERSON_ID = '" & Me.personID & "'"
            'Response.Write(SQL)
            'Response.End()

            rs = New SqlCommand(SQL, cn)
            rsReply = rs.ExecuteReader

            If rsReply.Read Then
                'fill data fields
                Me.FIRST_NAME.Text = rsReply("FIRST_NAME")
                Me.LAST_NAME.Text = rsReply("LAST_NAME")
                Me.EMAIL.Text = rsReply("EMAIL")
                Me.BIRTH_DATE.Text = Replace(rsReply("BIRTH_DATE"), "1/1/1900", "")
                Me.GENDER.SelectedValue = rsReply("GENDER")
                Me.PRIMARY_PHONE_NUM.Text = rsReply("PRIMARY_PHONE_NUM")
                Me.PHONE2.Text = rsReply("PHONE2")
                Me.PHONE3.Text = rsReply("PHONE3")
                Me.ADDRESS_LINE1.Text = rsReply("ADD_LINE1")
                Me.ADDRESS_LINE2.Text = rsReply("ADD_LINE2")
                Me.CITY.Text = rsReply("vCITY")
                Me.STATE.SelectedValue = rsReply("vSTATE")
                Me.ZIP.Text = rsReply("vZIP")
                Me.EMERGENCY_FIRST_NAME.Text = rsReply("vEMER_FNAME")
                Me.EMERGENCY_LAST_NAME.Text = rsReply("vEMER_LNAME")
                Me.EMERGENCY_PHONE.Text = rsReply("vEMER_PHONE")
                Me.SUPPRESS_EMAIL_IND.SelectedValue = rsReply("SUPPRESS_EMAIL_IND")
                Me.SUPPRESS_SNAIL_MAIL_IND.SelectedValue = rsReply("SUPPRESS_SNAIL_MAIL_IND")
                Me.SUPPRESS_LAST_NAME_IND.SelectedValue = rsReply("SUPPRESS_LAST_NAME_IND")
                Me.FIRST_CONTACT_ID.SelectedValue = rsReply("FIRST_CONTACT_ID")
                'Me.PHOTO_FILENAME.Text = rsReply("FIRST_NAME")
                Me.NAGVA_RATING.Text = rsReply("vNAGVA")
                Me.TSHIRT_SIZE.SelectedValue = rsReply("vTSHIRT")

            Else
                'redirect out of this area

            End If

            rsReply.Close()

            'see if there is an image for this user
            getImage()
        End If



    End Sub


    Public Sub getImage()
        'check for image file
        Dim imgPath As String = Request.ServerVariables("APPL_PHYSICAL_PATH") & "personnel_images\"
        Dim str As String
        Dim imgFound As Boolean = False
        For Each str In Directory.GetFiles(imgPath, Session("PERSON_ID").ToString + "_*")
            Me.imgFile.ImageUrl = "http://cgva.org/personnel_images/" + Path.GetFileName(str)
            Me.imgFile.Width = 200

            'Response.Write(Me.imgFile.ImageUrl & "<br />")
            imgFound = True
            Exit For
        Next

        If Not imgFound Then
            'Me.lblImage.Text = ""
            Me.lblImageFound.Text = "NO"
            Me.imgFile.Visible = False
        Else
            Me.lblImageFound.Text = ""
            Me.imgFile.Visible = True
        End If

    End Sub

    Public Sub removeImage()
        'check for image file
        Dim imgPath As String = Request.ServerVariables("APPL_PHYSICAL_PATH") & "personnel_images\"
        Dim str As String
        Dim imgFound As Boolean = False
        For Each str In Directory.GetFiles(imgPath, Session("PERSON_ID").ToString + "_*")
            'Response.Write(Request.ServerVariables("APPL_PHYSICAL_PATH") & "personnel_images\" & str)
            Try
                System.IO.File.Delete(str)
            Catch ex As Exception
                'Response.Write(ex.Message & "-" & str & "<br />")
            End Try
        Next

        'If Not imgFound Then
        'Me.lblImage.Text = ""
        'Me.imgFile.Visible = False
        'End If

    End Sub

    Public Function CheckNull(ByVal fieldValue As String) As String
        If fieldValue.Equals(DBNull.Value) Then
            Return ""
        Else
            Return fieldValue
        End If
    End Function

    Public Shared Function isEmail(ByVal inputEmail As String) As Boolean

        Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" & "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" & ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        Dim re As New Regex(strRegex)
        If re.IsMatch(inputEmail) Then
            Return (True)
        Else
            Return (False)
        End If

    End Function

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click

        If Not Page.IsValid Then
            Me.messageLabel.Text = "There is an error on the page."
            Exit Sub
        Else
            Me.messageLabel.Text = ""
        End If


        FNAME = Trim(Me.FIRST_NAME.Text)
        LNAME = Trim(Me.LAST_NAME.Text)
        strEmail = Trim(Me.EMAIL.Text)
        firstContactID = Me.FIRST_CONTACT_ID.SelectedValue
        strGender = Me.GENDER.SelectedValue
        strTshirt = Me.TSHIRT_SIZE.SelectedValue
        strNAGVA = Me.NAGVA_RATING.Text
        strPhone1 = Trim(Me.PRIMARY_PHONE_NUM.Text)
        strPhone2 = Trim(Me.PHONE2.Text)
        strPhone3 = Trim(Me.PHONE3.Text)
        strBirthDate = Trim(Me.BIRTH_DATE.Text)
        strADDRESS_LINE1 = Trim(Me.ADDRESS_LINE1.Text)
        strADDRESS_LINE2 = Trim(Me.ADDRESS_LINE2.Text)
        strCITY = Trim(Me.CITY.Text)
        strSTATE = Me.STATE.SelectedValue
        strZIP = Trim(Me.ZIP.Text)
        strEMERGENCY_FIRST_NAME = Trim(Me.EMERGENCY_FIRST_NAME.Text)
        strEMERGENCY_LAST_NAME = Trim(Me.EMERGENCY_LAST_NAME.Text)
        strEMERGENCY_PHONE = Trim(Me.EMERGENCY_PHONE.Text)
        strSUPPRESS_EMAIL_IND = Trim(Me.SUPPRESS_EMAIL_IND.SelectedValue)
        strSUPPRESS_SNAIL_MAIL_IND = Trim(Me.SUPPRESS_SNAIL_MAIL_IND.SelectedValue)
        strSUPPRESS_LAST_NAME_IND = Trim(Me.SUPPRESS_LAST_NAME_IND.SelectedValue)

        'strPassword1 = Trim(Me.PASSWORD1.Text)
        'strPassword2 = Trim(Me.PASSWORD2.Text)
        'securityQuestion = Me.SECURITY_QUESTION.SelectedValue
        'securityAnswer = Trim(Me.SECURITY_ANSWER.Text)


        If FNAME = "" Then
            Me.messageLabel.Text = "Please enter your first name."
            Exit Sub
        ElseIf LNAME = "" Then
            Me.messageLabel.Text = "Please enter your last name."
            Exit Sub
        ElseIf strEmail = "" Then
            Me.messageLabel.Text = "Please enter your email address."
            Exit Sub
        ElseIf Not isEmail(strEmail) Then
            Me.messageLabel.Text = "The email address entered is not valid."
            Exit Sub
        ElseIf firstContactID = 0 Then
            Me.messageLabel.Text = "Please enter how you first heard about CGVA."
            Exit Sub
        Else
            Dim cn As SqlConnection
            Dim rs As SqlCommand
            'Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr

            cn.Open()
            SQL = "UPDATE db_accessadmin.[PERSON_TBL] " _
                & "SET FIRST_NAME='" & Me.FNAME & "', " _
                & "LAST_NAME='" & Me.LNAME & "', " _
                & "EMAIL='" & Me.strEmail & "', " _
                & "FIRST_CONTACT_ID='" & Me.firstContactID.ToString & "', " _
                & "GENDER='" & Me.strGender & "', " _
                & "BIRTH_DATE='" & Me.strBirthDate & "', " _
                & "TSHIRT_SIZE='" & Me.strTshirt & "', " _
                & "NAGVA_RATING='" & Me.strNAGVA & "', " _
                & "PRIMARY_PHONE_NUM='" & Me.strPhone1 & "', " _
                & "[2ND_PHONE_NUM]='" & Me.strPhone2 & "', " _
                & "[3RD_PHONE_NUM]='" & Me.strPhone3 & "', " _
                & "ADDRESS_LINE1='" & Me.strADDRESS_LINE1 & "', " _
                & "ADDRESS_LINE2='" & Me.strADDRESS_LINE2 & "', " _
                & "CITY = '" & Me.strCITY & "', " _
                & "STATE='" & Me.strSTATE & "', " _
                & "ZIP='" & Me.strZIP & "', " _
                & "EMERGENCY_FIRST_NAME='" & Me.strEMERGENCY_FIRST_NAME & "', " _
                & "EMERGENCY_LAST_NAME='" & Me.strEMERGENCY_LAST_NAME & "', " _
                & "EMERGENCY_PHONE='" & Me.strEMERGENCY_PHONE & "', " _
                & "SUPPRESS_EMAIL_IND='" & Me.strSUPPRESS_EMAIL_IND & "', " _
                & "SUPPRESS_SNAIL_MAIL_IND='" & Me.strSUPPRESS_SNAIL_MAIL_IND & "', " _
                & "SUPPRESS_LAST_NAME_IND='" & Me.strSUPPRESS_LAST_NAME_IND & "' " _
                & "WHERE PERSON_ID='" & Me.personID & "'"
            'Response.Write(SQL)
            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()

            Try
                'upload new image
                Dim sFileDir As String = Request.ServerVariables("APPL_PHYSICAL_PATH") & "personnel_images\"
                'picture to upload?
                If (Not PHOTO_NAME.PostedFile Is Nothing) _
                   And (PHOTO_NAME.PostedFile.ContentLength > 0) Then

                    'get rid of old image first
                    removeImage()

                    'determine file name
                    Dim sFileName As String = _
                       System.IO.Path.GetFileName(PHOTO_NAME.PostedFile.FileName)
                    System.IO.Path.GetFileName(PHOTO_NAME.PostedFile.FileName)
                    'save file on disk
                    PHOTO_NAME.PostedFile.SaveAs(sFileDir + Session("PERSON_ID").ToString + "_" + sFileName)

                    Me.lblImageFound.Text = "Yes"
                    Me.imgFile.Visible = False
                    'show new image
                    'getImage()
                End If

            Catch ex As Exception
                Me.messageLabel.Text = ex.StackTrace + "The image file you attempted to upload was not valid.<br />Please contact support@cgva.org for assistance with this issue."
                Exit Sub
            End Try

            Me.messageLabel.Text = "<font class='cfontSuccess10'>Your profile information has been updated successfully.</font>"

        End If


    End Sub
End Class
