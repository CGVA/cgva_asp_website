Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class MyEvents2
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

    Private strRating As String
    Private strRatingDate As String

    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Private eventArray As New ArrayList


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString
        Else
            personID = "0"
        End If

        If personID = "0" Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If


        'If Not IsPostBack Then

        Me.messageRow.Visible = False

        If Not InStr(Request.ServerVariables("HTTP_REFERER"), "MyEvents.aspx", CompareMethod.Text) > 0 Then
            'If Me.vS2 <> "on" And Me.vS4 <> "on" And Me.vS6 <> "on" Then
            Response.Redirect("MyEvents.aspx")
        End If

        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsReply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        SQL = "SELECT EVENT_CD FROM EVENT_TBL " _
            & "WHERE EVENT_CD IN " _
            & "(SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y')"

        rs = New SqlCommand(SQL, cn)
        rsReply = rs.ExecuteReader

        If rsReply.HasRows Then
            While rsReply.Read
                'Response.Write(rsReply("EVENT_CD"))
                'Response.End()
                Me.eventArray.Add(rsReply("EVENT_CD"))
            End While

            rsReply.Close()

        Else
            'shouldnt get here, but redirect back to MyEvents if we do
            cn.Close()
            Response.Redirect("MyEvents.aspx")
            Response.End()
        End If

        'create hidden fields for any chosen events, maintaining the values for next page
        Dim tempString As String = ""

        For Each s As String In Me.eventArray

            Dim vHiddenField As New HiddenField
            vHiddenField.ID = s
            vHiddenField.Value = ""

            For i = 0 To Request.Form.Count - 1
                'Response.Write(Request.Form.AllKeys(i) & ":")
                'Response.Write(Request.Form(i) & "<br />")
                tempString = Request.Form.AllKeys(i)

                If s = tempString Then
                    'Response.Write("equal" + s + "<br />")
                    vHiddenField.Value = Request.Form(i)
                End If

            Next

            Me.Form1.Controls.Add(vHiddenField)

        Next

        'Response.End()

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

        rs.CommandText = SQL
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

        'get most recent rating
        SQL = "SELECT IsNull(RATING_SCORE,'unrated') as 'RATING',CONVERT(varchar,r.EFF_DATE,101) as 'EFF_DATE' " _
        & "FROM	RATINGS_TBL r " _
        & "WHERE r.PERSON_ID = '" & Session("PERSON_ID").ToString & "' " _
        & "AND r.EFF_DATE = " _
        & "(SELECT MAX(EFF_DATE) FROM RATINGS_TBL WHERE PERSON_ID = r.PERSON_ID)"
        rs = New SqlCommand(SQL, cn)
        rsReply = rs.ExecuteReader

        If rsReply.Read Then
            strRating = rsReply("RATING").ToString
            strRatingDate = rsReply("EFF_DATE").ToString

            Me.ratingLabel.Text = "Your current rating is <b>" + strRating + "</b>. You were last rated on " + strRatingDate + ". If you like to be re-evaluated, please inform the registration committee."
        Else
            strRating = "unrated"
            Me.ratingLabel.Text = "You are currently <b>" + strRating + "</b>. Please inform the registration committee that you need to be rated."
        End If

        rsReply.Close()
        cn.Close()

        'End If

        'Response.Write("League:" + League.Value.ToString + ", SB:" + SB.Value.ToString)
        'Me.submitButton.Enabled = False
        'Me.verifyInfo.Attributes.Add("onclick", "verifyInfo_CheckedChanged")
    End Sub

    Protected Sub verifyInfo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles verifyInfo.CheckedChanged
        'Response.Write("HERE")
        If Me.verifyInfo.Checked Then
            Me.submitButton.Enabled = True
        Else
            Me.submitButton.Enabled = False
        End If
    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        If Not Me.verifyInfo.Checked Then
            Me.messageRow.Visible = True
            Me.messageLabel.Text = "<font class=""cfontError10"">Please verify all of your personal information, and check the verification box at the bottom of the screen to continue.</font>"
        Else
            'Context.Items.Add("League", League.Value)
            'Context.Items.Add("SB", SB.Value)
            Server.Transfer("MyEvents3.aspx", True)
            'Server.Transfer("MyEvents5.aspx", True)
        End If

    End Sub
End Class
