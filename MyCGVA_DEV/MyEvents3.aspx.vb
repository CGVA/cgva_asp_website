Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class MyEvents3
    Inherits System.Web.UI.Page

    Private vVerifyInfo As String
    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Public personID As String = "0"
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

        'If Not InStr(Request.ServerVariables("HTTP_REFERER"), "MyEvents2.aspx", CompareMethod.Text) > 0 Then
        'Response.Redirect("MyEvents.aspx")
        'End If

        'If Not IsPostBack Then
        vVerifyInfo = Request.Form("verifyInfo")

        If Me.vVerifyInfo <> "on" Then
            Response.Redirect("MyEvents.aspx")
        Else
            Dim vHiddenField As New HiddenField
            vHiddenField.ID = "verifyInfo"
            vHiddenField.Value = "on"
            Me.Form1.Controls.Add(vHiddenField)

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

        'Response.Write("League3:" + vLeague.ToString + ", SB3:" + vSB.ToString)

        'If Not vS2 = "on" Then
        'Me.LeagueWaiver.Text = ""
        'End If

#If comment Then
        '5/1/2011 come back to this later if needed, dynamic pulls of waiver information
        SQL = "SELECT LOCATION_CD, " _
            & "EVENT_SHORT_DESC, " _
            & "EVENT_LONG_DESC, " _
            & "EVENT_START_DATE, " _
            & "EVENT_END_DATE " _
            & "FROM EVENT_TBL " _
            & "WHERE EVENT_CD IN ('2011S2')"
        rs.CommandText = SQL
        rsReply = rs.ExecuteReader

        If rsReply.Read Then
            ''get data - TO DO
        End If

        rsReply.Close()
#End If
        cn.Close()

        'End If

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        If Not Me.verifyWaiverInfo.Checked Then
            Me.messageLabel.Text = "Please read the waiver information, and electronically sign the waiver at the bottom of the screen to continue with registration."
        Else
            'JPC 4/26/11 bypass this next page for summer league
            'Server.Transfer("MyEvents4.aspx", True)
            Server.Transfer("MyEvents5.aspx", True)
        End If

    End Sub
End Class
