Imports System.Data.SqlClient
Imports System.Net.Mail


Partial Class MyEvents5
    Inherits System.Web.UI.Page

    Private vVerifyInfo As String
    Private vVerifyWaiverInfo As String
    Private vS2, vS4, vS6 As String
    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Public personID As String = "0"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")
    'use this value if DDL will have multiple values 
    'Private intAmount As Integer = 0
    Private intAmount As Integer = 75

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
        Me.vVerifyInfo = Request.Form("verifyInfo")

        '11/5/11 add here because we took out the waiver
        If Me.vVerifyInfo <> "on" Then
            Response.Redirect("MyEvents.aspx")
        Else
            Dim vHiddenField As New HiddenField
            vHiddenField.ID = "verifyInfo"
            vHiddenField.Value = "on"
            Me.form1.Controls.Add(vHiddenField)

            'Dim vHiddenField2 As New HiddenField
            'vHiddenField.ID = "verifyWaiverInfo"
            'vHiddenField.Value = vVerifyWaiverInfo
            'Me.form1.Controls.Add(vHiddenField2)

        End If

        '11/5/11 no waiver now
        'Me.vVerifyWaiverInfo = Request.Form("verifyWaiverInfo")
        'If Me.vVerifyInfo <> "on" And Me.vVerifyWaiverInfo <> "on" Then
        'Response.Redirect("MyEvents.aspx")
        'End If

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

        'hardcoded for summer leagues
        'Dim s4, s6 As Boolean
        's4 = False
        's6 = False

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
                    'TO DO create DDL value/amount for this selected leagues

                    'hardcode for double summer league
                    'If s = "2012S4" And vHiddenField.Value = "on" Then
                    's4 = True
                    'ElseIf s = "2012S6" And vHiddenField.Value = "on" Then
                    's6 = True
                    'End If
                End If

            Next

            Me.form1.Controls.Add(vHiddenField)

        Next


        ' FALL LEAGUE 2013
        leagueChoiceDropDownList.Items.Insert(0, New ListItem("$75.00 - 2014 Spring League Fee", "75"))
        Session("Amount") = intAmount.ToString
        Me.amount.Value = Session("Amount")

        'for now hardcode this
        'VOLLEYPALOOZA 2013
        'If Not IsPostBack Then
        '    intAmount = 35
        '    leagueChoiceDropDownList.Items.Insert(0, New ListItem("$35.00 - 2013 Volleypalooza Fee", "35"))
        '    leagueChoiceDropDownList.Items.Insert(1, New ListItem("$140.00 - 2013 Volleypalooza Team Fee - Four Players", "140"))
        '    leagueChoiceDropDownList.Items.Insert(2, New ListItem("$175.00 - 2013 Volleypalooza Team Fee - Five Players", "175"))
        '    Me.leagueChoiceDropDownList.Enabled = True
        '    Session("Amount") = intAmount.ToString
        '    Me.amount.Value = Session("Amount")
        'End If


        'TO DO create, for now hardcode this
        'If s4 = True And s6 = True Then
        'leagueChoiceDropDownList.Items.Insert(0, New ListItem("$150.00 - 2012 Summer League Fees", "2012 Summer League Fees"))
        'intAmount = 150
        'Me.leagueChoiceDropDownList.SelectedIndex = 0
        'ElseIf s4 = True Then
        'leagueChoiceDropDownList.Items.Insert(0, New ListItem("$85.00 - 2012 Summer League Fees", "2012 Summer League Fees"))
        'intAmount = 85
        'Me.leagueChoiceDropDownList.SelectedIndex = 0
        'ElseIf s6 = True Then
        'leagueChoiceDropDownList.Items.Insert(0, New ListItem("$65.00 - 2012 Summer League Fees", "2012 Summer League Fees"))
        'intAmount = 65
        'Me.leagueChoiceDropDownList.SelectedIndex = 0
        'End If


        'JPC 12/6/09 used for fall leagues 2009
        'Me.CAPTAIN_IND.Value = Request.Form("CAPTAIN_IND")
        'Me.CAPTAIN_INTEREST.Value = Request.Form("CAPTAIN_INTEREST")
        'Me.CROSS_DIVISION_IND.Value = Request.Form("CROSS_DIVISION_IND")
        'Session("CAPTAIN_IND") = Me.CAPTAIN_IND.Value
        'Session("CAPTAIN_INTEREST") = Me.CAPTAIN_INTEREST.Value
        'Session("CROSS_DIVISION_IND") = Me.CROSS_DIVISION_IND.Value

        'JPC 12/6/09 used for spring leagues 2010
        'bypassed for summer 2011
        'Me.CAPTAIN_IND_League.Value = Request.Form("CAPTAIN_IND_League")
        'Session("CAPTAIN_IND_League") = Me.CAPTAIN_IND_League.Value

        'going to default everyone to captain for summer
        'Me.CAPTAIN_IND_League.Value = "Y"
        'Session("CAPTAIN_IND_League") = Me.CAPTAIN_IND_League.Value

        'JPC 12/6/09 used for fall leagues 2009
        'Me.Q1.Value = Request.Form("Q1")
        'Me.Q2.Value = Request.Form("Q2")
        'Me.Q3.Value = Request.Form("Q3")
        'Me.Q4.Value = Request.Form("Q4")
        'Session("Q1") = Me.Q1.Value
        'Session("Q2") = Me.Q2.Value
        'Session("Q3") = Me.Q3.Value
        'Session("Q4") = Me.Q4.Value

        'bypassed for summer 2011, no summer rating requests
        'Me.Q5.Value = Request.Form("Q5")
        'Session("Q5") = Me.Q5.Value
        'Me.Q5.Value = "N"
        'Session("Q5") = "N"

        'Response.Write(Session("League") + "<br />")
        'Response.Write(Session("SB") + "<br />")
        'Response.Write(Session("verifyInfo") + "<br />")
        'Response.Write(Session("verifyWaiverInfo") + "<br />")
        'Response.Write(Session("amount") + "<br />")
        'Response.Write(Me.CAPTAIN_IND.Value + "<br />")
        'Response.Write(Me.CAPTAIN_INTEREST.Value + "<br />")
        'Response.Write(Me.CROSS_DIVISION_IND.Value + "<br />")
        'Response.Write(Me.Q1.Value + "<br />")
        'Response.Write(Me.Q2.Value + "<br />")
        'Response.Write(Me.Q3.Value + "<br />")
        'Response.Write(Me.Q4.Value + "<br />")
        'Response.Write(Me.Q5.Value + "<br />")

        'End If

    End Sub

    Protected Sub paypalImageButton_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles paypalImageButton.Click

        'Response.Redirect("PayPal.aspx")
        Server.Transfer("PayPal.aspx", True)
    End Sub

    'USE FOR VOLLEYPALOOZA2013
    Protected Sub leagueChoiceDropDownList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles leagueChoiceDropDownList.SelectedIndexChanged
        'Response.Write(Me.leagueChoiceDropDownList.SelectedIndex.ToString + "<br />")
        Try
            Select Case Me.leagueChoiceDropDownList.SelectedIndex
                Case 0
                    Session("Amount") = "35"
                    'Session("Amount") += "Case 0"
                    'Me.lblText.Text = Session("Amount")
                    Me.amount.Value = Session("Amount")
                Case 1
                    Session("Amount") = "140"
                    'Session("Amount") += "Case 1"
                    'Me.lblText.Text = Session("Amount")
                    Me.amount.Value = Session("Amount")
                Case 2
                    Session("Amount") = "175"
                    'Session("Amount") += "Case 2"
                    'Me.lblText.Text = Session("Amount")
                    Me.amount.Value = Session("Amount")
                Case Else
                    Session("Amount") = "35"
                    'Session("Amount") += "Case 3"
                    'Me.lblText.Text = Session("Amount")
                    Me.amount.Value = Session("Amount")
            End Select
        Catch ex As Exception
            Session("Amount") = "35"
            Me.amount.Value = Session("Amount")
        End Try
    End Sub
End Class
