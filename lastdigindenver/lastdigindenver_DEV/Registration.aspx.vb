Imports System.Data.SqlClient

Public Class Registration
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.successLabel.Text = ""

        If Not IsPostBack Then
            If Request.QueryString("success") = "true" Then
                Me.successLabel.Text = "Congratulations!<br />Once your payment has been verified, your transaction will be complete."
                Me.successLabel.ForeColor = Drawing.Color.Green
            ElseIf Request.QueryString("success") = "false" Then
                Me.successLabel.Text = "Your registration request has been canceled."
                Me.successLabel.ForeColor = Drawing.Color.Red
            End If

            Me.errorLabelTeam.Text = ""
            Me.errorLabelAdditionalPlayers.Text = ""

            Dim SQL As String = ""
            Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            SQL = "SELECT CONVERT(VARCHAR(10), START_DTG, 101)  as 'START_DTG', " _
                    & "CONVERT(VARCHAR(10), END_DTG, 101)  as 'END_DTG', " _
                    & "CONVERT(VARCHAR,TEAM_FEE_PAYPAL,0) as 'TEAM_FEE_PAYPAL', " _
                    & "CONVERT(VARCHAR,IND_FEE_PAYPAL,0) as 'IND_FEE_PAYPAL', " _
                    & "CONVERT(VARCHAR,TEAM_FEE_MAIL,0) as 'TEAM_FEE_MAIL', " _
                    & "CONVERT(VARCHAR,IND_FEE_MAIL,0) as 'IND_FEE_MAIL' " _
                    & "from LASTDIG_REG_DTG_TBL WHERE getdate() between START_DTG AND END_DTG"

            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            'Response.Write(SQL)
            If rsreply.Read Then
                Me.openPanel.Visible = True
                Me.closedPanel.Visible = False

                Me.teamFeeLabel.Text = rsreply("TEAM_FEE_PAYPAL").ToString
                Me.registrationDateLabel.Text = rsreply("END_DTG").ToString
                Me.additionalPayerFeeLabel.Text = rsreply("IND_FEE_PAYPAL").ToString

                Me.additionalPlayersDDL.Items.Add(New ListItem("1 ($" + (1 * rsreply("IND_FEE_PAYPAL")).ToString + ")", rsreply("IND_FEE_PAYPAL").ToString))
                Me.additionalPlayersDDL.Items.Add(New ListItem("2 ($" + (2 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (2 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL.Items.Add(New ListItem("3 ($" + (3 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (3 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL.Items.Add(New ListItem("4 ($" + (4 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (4 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL.Items.Add(New ListItem("5 ($" + (5 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (5 * rsreply("IND_FEE_PAYPAL")).ToString))

                Me.teamFee.Value = rsreply("TEAM_FEE_PAYPAL").ToString

                Me.additionalPlayersDDL2.Items.Add(New ListItem("1 ($" + (1 * rsreply("IND_FEE_PAYPAL")).ToString + ")", rsreply("IND_FEE_PAYPAL").ToString))
                Me.additionalPlayersDDL2.Items.Add(New ListItem("2 ($" + (2 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (2 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL2.Items.Add(New ListItem("3 ($" + (3 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (3 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL2.Items.Add(New ListItem("4 ($" + (4 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (4 * rsreply("IND_FEE_PAYPAL")).ToString))
                Me.additionalPlayersDDL2.Items.Add(New ListItem("5 ($" + (5 * rsreply("IND_FEE_PAYPAL")).ToString + ")", (5 * rsreply("IND_FEE_PAYPAL")).ToString))

                rsreply.Close()

                SQL = "SELECT DISTINCT(TEAM_NAME + ' - ' + DIVISION) as 'TEAM_NAME' " _
                       & "FROM LASTDIG_TEAM_TBL WHERE ARCHIVED_DTG IS NULL " _
                       & "AND ACTIVE = 'Y'"
                'no longer need this
                '                       & "FROM LASTDIG_TEAM_TBL WHERE RECONCILE_DTG IS NOT NULL " _

                rs = New SqlCommand(SQL, cn)
                rsreply = rs.ExecuteReader

                While rsreply.Read
                    Me.teamNameDDL.Items.Add(New ListItem(rsreply("TEAM_NAME").ToString, rsreply("TEAM_NAME").ToString))
                End While

                rsreply.Close()
                cn.Close()
            Else
                cn.Close()
                Me.openPanel.Visible = False
                Me.closedPanel.Visible = True
            End If


        End If


        'Response.Write("HERE")
        'Response.End

    End Sub

    Private Function validateTeam(ByVal teamName As String, ByVal division As String) As Boolean
        Me.errorLabelTeam.Text = ""
        Me.errorLabelAdditionalPlayers.Text = ""

        If Trim(Me.teamName.Text) = "" Then
            Me.errorLabelTeam.Text = "Please enter a team name to continue with registration."
            Me.teamName.Focus()
            validateTeam = False
        ElseIf Me.divisionDDL.SelectedIndex = 0 Then
            Me.errorLabelTeam.Text = "Please select your division to continue with registration."
            Me.divisionDDL.Focus()
            validateTeam = False
        Else
            Dim SQL As String = ""
            Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()
            'VERIFY TEAM NAME DOESN'T ALREADY EXIST FOR THIS DIVISION
            SQL = "SELECT COUNT(*) FROM LASTDIG_TEAM_TBL " _
                    & "WHERE TEAM_NAME='" & teamName & "' " _
                    & "AND DIVISION = '" & division & "'" _
                    & "AND ACTIVE = 'Y' AND ARCHIVED_DTG IS NULL"

            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            rsreply.Read()
            If rsreply(0) > 0 Then
                Me.errorLabelTeam.Text = "The team name you entered already exists in the selected division.<br />Please enter a new name to continue with registration."
                Me.teamName.Focus()
                validateTeam = False
            Else
                validateTeam = True
            End If

            rsreply.Close()
            cn.Close()

        End If

    End Function

    Private Function validateAdditionalPlayer() As Boolean
        Me.errorLabelTeam.Text = ""
        Me.errorLabelAdditionalPlayers.Text = ""

        If Me.teamNameDDL.SelectedIndex = 0 Then
            Me.errorLabelAdditionalPlayers.Text = "Please select your team name/division to register additional players."
            Me.teamNameDDL.Focus()
            validateAdditionalPlayer = False
        ElseIf Me.additionalPlayersDDL2.SelectedIndex = 0 Then
            Me.errorLabelAdditionalPlayers.Text = "Please select how many additional players you would like to add to your team."
            Me.additionalPlayersDDL2.Focus()
            validateAdditionalPlayer = False
        Else
            validateAdditionalPlayer = True
        End If
    End Function

    Protected Sub paypalImageButton2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles paypalImageButton2.Click

        If validateAdditionalPlayer() Then
            Dim teamName As String = ""
            Dim division As String = ""
            Dim teamFee As String = ""
            Dim additionalPlayers As String = ""
            Dim totalFee As Integer = 0
            Dim transactionID As Integer = 0
            Dim teamNameTemp As String = ""
            Dim teamNameArray As Array
            teamNameTemp = Replace(Request("teamNameDDL"), "'", "''")
            teamNameArray = Split(teamNameTemp, " - ")
            teamName = teamNameArray(0)
            division = teamNameArray(1)
            additionalPlayers = Request("additionalPlayersDDL2")
            totalFee = CInt(additionalPlayers)

            Session("totalFee") = totalFee.ToString
            Session("custom") = (Replace(teamName, """", "") + "~" + division + "~" + totalFee.ToString).ToString

            'Response.Redirect("http://cgva.org/lastdigindenver/PayPalLASTDIG.aspx")
            'Response.End()
            Server.Transfer("PayPalLASTDIG.aspx", False)
        End If

    End Sub

    Protected Sub paypalImageButton_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles paypalImageButton.Click

        Dim teamName As String = ""
        Dim division As String = ""
        Dim teamFee As String = ""
        Dim additionalPlayers As String = ""
        Dim totalFee As Integer = 0
        Dim transactionID As Integer = 0

        teamName = Replace(Request("teamName"), "'", "''")
        teamName = Replace(teamName, """", "")
        division = Request("divisionDDL")
        teamFee = Request("teamFee")
        additionalPlayers = Request("additionalPlayersDDL")
        totalFee = CInt(teamFee) + CInt(additionalPlayers)

        If validateTeam(teamName, division) Then

            Session("totalFee") = totalFee.ToString
            Session("custom") = (Replace(teamName, """", "") + "~" + division + "~" + totalFee.ToString).ToString

            'Response.Redirect("http://cgva.org/lastdigindenver/lastdigindenver_DEV/PayPalLASTDIG.aspx")
            'Response.End()
            Server.Transfer("PayPalLASTDIG.aspx", False)

        End If


    End Sub

End Class
