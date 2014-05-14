Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class MyEvents4
    Inherits System.Web.UI.Page

    Private vVerifyInfo As String
    Private vVerifyWaiverInfo As String
    Private vDU2 As String
    Private vDU4 As String

    Private strRating As String
    Private strRatingDate As String

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Public personID As String = "0"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '4/1/2010 JPC this page not needed for volleypalooza, so let's go right to page 5
        Server.Transfer("MyEvents5.aspx", True)


        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString
        Else
            personID = "0"
        End If

        If personID = "0" Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If


        If Not IsPostBack Then
            Me.vVerifyInfo = Request.Form("verifyInfo")
            Me.vVerifyWaiverInfo = Request.Form("verifyWaiverInfo")

            If Me.vVerifyInfo <> "on" And Me.vVerifyWaiverInfo <> "on" Then
                Response.Redirect("MyEvents.aspx")
            End If
#If Comment Then

            'vDU2 = Request.Form("DU2")
            vDU4 = Request.Form("DU4")
            Me.DU2.Value = vDU2
            Me.DU4.Value = vDU4
            Me.verifyInfo.Value = vVerifyInfo
            Me.verifyWaiverInfo.Value = vVerifyWaiverInfo

            'Response.Write("DU4:" + vDU2.ToString + ", SB4:" + vDU4.ToString)

            'JPC 12/15/09 not needed with spring team creation code
            'If vDU4 = "on" Then
            'Me.SBPanel.Visible = True
            'Else
            Me.DU4Panel.Visible = False
            'End If

            'If vDU2 = "on" Then
            'Me.DUPanel.Visible = True
            'Else
            Me.DU2Panel.Visible = False
            'End If

            Me.skillsDevelopmentPanel.Visible = False

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

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

                Me.ratingLabel.Text = "Your current rating is <b>" + strRating + "</b>. You were last rated on " + strRatingDate + ". Would you like to be re-evaluated?"
            Else
                strRating = "unrated"
                Me.ratingLabel.Text = "You are currently <b>" + strRating + "</b>. The ratings committee will contact you."
                Me.Q5.SelectedValue = "Y"
                Me.Q5.Enabled = False
                Me.Q5.Visible = True
            End If

            rsReply.Close()
            cn.Close()
#End If

        End If

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        'If Not Me.verifyInfo.Checked Then
        'Me.messageLabel.Text = "Please read the waiver information, and electronically sign the waiver at the bottom of the screen to continue with registration."
        'Else
        'Context.Items.Add("DU2", DU2.Value)
        'Context.Items.Add("DU4", DU4.Value)
        Server.Transfer("MyEvents5.aspx", True)
        'End If

    End Sub
End Class
