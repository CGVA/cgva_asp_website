Imports System.Data.SqlClient

Public Class Registration
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then
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
                       & "FROM LASTDIG_TEAM_TBL WHERE RECONCILE_DTG IS NOT NULL"

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

    Private Function validateTeam() As Boolean
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
            validateTeam = True
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

        End If

    End Sub

    Protected Sub paypalImageButton_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles paypalImageButton.Click

        If validateTeam() Then

            Dim teamName As String = ""
            Dim division As String = ""
            Dim teamFee As String = ""
            Dim additionalPlayers As String = ""
            Dim totalFee As Integer = 0
            Dim newID As Integer = 0

            Dim SQL As String = ""
            Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            If Request("registrationType") = "TEAM" Then
                teamName = Replace(Request("teamName"), "'", "''")
                division = Request("division")
                teamFee = Request("teamFee")
                additionalPlayers = Request("additionalPlayers")
                totalFee = CInt(teamFee) + CInt(additionalPlayers)

                'VERIFY TEAM NAME DOESN'T ALREADY EXIST FOR THIS DIVISION
                SQL = "SELECT COUNT(*) FROM LASTDIG_TEAM_TBL " _
                        & "WHERE TEAM_NAME='" & teamName & "' " _
                        & "AND DIVISION = '" & division & "'"

                rs = New SqlCommand(SQL, cn)
                rsreply = rs.ExecuteReader
                rsreply.Read()
                If rsreply(0) > 0 Then
                    Session("Err") = "The team name you entered already exists in the selected division. Please enter a new name to continue with registration."
                    cn.Close()
                    Response.Redirect("paypal.asp")
                End If
                rsreply.Close()

                SQL = "Set Nocount on " _
                    & "INSERT INTO LASTDIG_TEAM_TBL(TEAM_NAME,DIVISION,FEE_PAID) " _
                    & "SELECT '" & teamName & "', " _
                    & "'" & division & "', " _
                    & "'" & totalFee & "' " _
                    & "SELECT @@IDENTITY  " _
                    & "Set Nocount OFF"
                '& "'" & totalFee & "'"

                'Response.Write(SQL)
                rs = New SqlCommand(SQL, cn)
                rsreply = rs.ExecuteReader
                rsreply.Read()
                newID = rsreply(0)
                'Response.Write(newID)
                'Response.End
                rsreply.Close()
                cn.Close()
            ElseIf Request("registrationType") = "PLAYER" Then
                Dim teamNameTemp As String = ""
                Dim teamNameArray As Array
                teamNameTemp = Replace(Request("teamName"), "'", "''")
                teamNameArray = Split(teamNameTemp, " - ")
                teamName = teamNameArray(0)
                division = teamNameArray(1)
                additionalPlayers = Request("additionalPlayers")
                totalFee = CInt(additionalPlayers)
                SQL = "Set Nocount on " _
                    & "INSERT INTO LASTDIG_TEAM_TBL(TEAM_NAME,DIVISION,FEE_PAID) " _
                    & "SELECT '" & teamName & "', " _
                    & "'" & division & "', " _
                    & "'" & totalFee & "' " _
                    & "SELECT @@IDENTITY " _
                    & "Set Nocount OFF"

                'Response.Write(SQL)
                rs = New SqlCommand(SQL, cn)
                rsreply = rs.ExecuteReader
                rsreply.Read()
                newID = rsreply(0)
                rsreply.Close()
                cn.Close()

                Response.Write(newID)
                Response.End()
            Else
                'ERROR
            End If

            Response.Redirect("http://cgva.org/MyCGVA/PayPalLASTDIG.aspx")
            Response.End()

        End If


    End Sub

End Class
