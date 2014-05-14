Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class TeamMaintenance
    Inherits System.Web.UI.Page

    Private personID As Integer = 0
    Private teamID As Integer = 0
    Private EVENT_CODE As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Response.Write("LOAD")
        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString()
        Else
            personID = 0
        End If

        If personID = 0 Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If

        Me.teamID = Request("ID")
        Me.EVENT_CODE = Request("EVENT_CD")
        Me.ID.Value = Request("ID")
        Me.EVENT_CD.Value = Request("EVENT_CD")

        If Not IsPostBack Then
            'Response.Write("NOT POSTBACK")
            'verify the person accessing this team is indeed the captain
            If validateCaptain() Then
                'loadUnassignedPlayersDDL()
                getTeamRoster()
            Else
                Me.messageLabel.Text = "You do not have access to view this team's roster."
                Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
                Me.submitButton.Visible = False
                Me.teamTable.Visible = False
                Exit Sub
            End If

        Else
            'Response.Write("POSTBACK")
            If Not validateCaptain() Then
                Me.messageLabel.Text = "You do not have access to view this team's roster."
                Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
                Me.submitButton.Visible = False
                Me.teamTable.Visible = False
                Exit Sub
            End If

        End If

    End Sub

    Protected Function validateCaptain() As Boolean
        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        SQL = "SELECT count(*) FROM TEAM_MEMBER_TBL " _
            & "WHERE PERSON_ID='" & personID.ToString & "' " _
            & "AND TEAM_ID='" & teamID.ToString & "' " _
            & "AND CAPTAIN_IND = 'Y'"

        'Response.Write(SQL)
        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader
        rsreply.Read()

        If rsreply(0) > 0 Then
            validateCaptain = True
        Else
            validateCaptain = False
        End If
        rsreply.Close()
        cn.Close()

    End Function

    Private Function loadUnassignedPlayersDDL() As DropDownList
        Dim unassignedPlayerDDL As New DropDownList
        Dim unassignedPlayerSQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        unassignedPlayerSQL = "SELECT a.PERSON_ID, b.FIRST_NAME, b.LAST_NAME " _
                        & "FROM REGISTRATION_TBL a " _
                        & "LEFT JOIN db_accessadmin.PERSON_TBL b ON a.PERSON_ID = b.PERSON_ID " _
                        & "WHERE a.EVENT_CD = '" & EVENT_CODE & "' " _
                        & "AND a.REGISTRATION_IND = 'Y' " _
                        & "AND a.PERSON_ID NOT IN (SELECT PERSON_ID FROM TEAM_MEMBER_TBL tm " _
                        & "							LEFT JOIN TEAM_TBL t ON tm.TEAM_ID = t.TEAM_ID " _
                        & "							LEFT JOIN DIVISION_TBL d ON t.DIVISION_ID = d.DIVISION_ID " _
                        & "							WHERE d.EVENT_CD = '" & EVENT_CODE & "') " _
                        & "ORDER BY UPPER(b.LAST_NAME), UPPER(b.FIRST_NAME)"


        unassignedPlayerDDL.Items.Add(New ListItem("SELECT", "0"))
        rs = New SqlCommand(unassignedPlayerSQL, cn)
        rsreply = rs.ExecuteReader



        While rsreply.Read
            unassignedPlayerDDL.Items.Add(New ListItem(rsreply(2) & ", " & rsreply(1), rsreply(0)))
        End While

        rsreply.Close()
        cn.Close()
        loadUnassignedPlayersDDL = unassignedPlayerDDL
    End Function

    Protected Sub getTeamRoster()
        Dim headerSQL, teamSQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs, rsTeam As SqlCommand
        Dim rsreply, rsTeamreply As SqlDataReader
        Dim ODD_ROW As Boolean = False
        Dim BGCOLOR As System.Drawing.Color
        Dim playerCount As Integer = 0

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        headerSQL = "SELECT e.EVENT_SHORT_DESC, " _
                   & "T.TEAM_NAME " _
                   & "FROM TEAM_TBL t " _
                   & "LEFT JOIN DIVISION_TBL d ON t.DIVISION_ID = d.DIVISION_ID " _
                   & "LEFT JOIN EVENT_TBL e ON d.EVENT_CD = e.EVENT_CD " _
                   & "WHERE e.EVENT_CD = '" & EVENT_CODE & "' " _
                   & "AND t.TEAM_ID = '" & teamID & "' " _
                   & "ORDER BY UPPER(EVENT_SHORT_DESC), DIVISION_DESC, TEAM_CD"
        'Response.Write(headerSQL)
        rs = New SqlCommand(headerSQL, cn)
        rsreply = rs.ExecuteReader
        If rsreply.Read Then
            Me.messageLabel.Text = "Please add/remove players on your team. Then click 'Save My Team'. " _
                                    & "<p /><u>" & rsreply(0) & "-" & rsreply(1) & "</u>"
        End If
        rsreply.Close()

        teamSQL = "SELECT t.PERSON_ID, p.FIRST_NAME, p.LAST_NAME " _
                & "FROM TEAM_MEMBER_TBL t " _
                & "LEFT JOIN db_accessadmin.PERSON_TBL p ON t.PERSON_ID = p.PERSON_ID " _
                & "WHERE TEAM_ID = '" & Me.teamID.ToString & "' " _
                & "ORDER BY t.CAPTAIN_IND DESC, UPPER(p.LAST_NAME), UPPER(p.FIRST_NAME)"
        'Response.Write(teamSQL)

        rsTeam = New SqlCommand(teamSQL, cn)
        rsTeamreply = rsTeam.ExecuteReader
        While rsTeamreply.Read
            playerCount += 1
            ODD_ROW = Not ODD_ROW

            If ODD_ROW Then
                BGCOLOR = System.Drawing.Color.White
            Else
                BGCOLOR = System.Drawing.Color.LightGray
            End If

            Dim tRow As New TableRow
            Dim tCounterCell As New TableCell
            Dim tPlayerNameCell As New TableCell

            Dim tCounterLabel As New Label
            Dim tPlayerDDL As New DropDownList

            tCounterLabel.Text = playerCount.ToString + "."
            tPlayerDDL.ID = "PERSON_ID" + playerCount.ToString
            tPlayerDDL.Items.Add(New ListItem("SELECT", "0"))
            tPlayerDDL.Items.Add(New ListItem(rsTeamreply(2) & ", " & rsTeamreply(1), rsTeamreply(0)))
            tPlayerDDL.SelectedIndex = 1
            tPlayerDDL.Width = 200

            'dont allow captain to remove himself from the team
            If rsTeamreply(0) = personID Then
                tPlayerDDL.Enabled = False
            End If

            tCounterCell.Controls.Add(tCounterLabel)
            tCounterCell.HorizontalAlign = HorizontalAlign.Right
            tCounterCell.CssClass = "cfont10"
            tPlayerNameCell.Controls.Add(tPlayerDDL)
            tPlayerNameCell.HorizontalAlign = HorizontalAlign.Left

            tRow.Cells.Add(tCounterCell)
            tRow.Cells.Add(tPlayerNameCell)
            tRow.BackColor = BGCOLOR
            Me.teamTable.Rows.Add(tRow)

        End While
        rsTeamreply.Close()
        cn.Close()

        For counter As Integer = playerCount To 9       'regular leagues have max of 10 players
            'For counter As Integer = playerCount To 1       'draft league with 2 core members
            playerCount += 1
            ODD_ROW = Not ODD_ROW

            If ODD_ROW Then
                BGCOLOR = System.Drawing.Color.White
            Else
                BGCOLOR = System.Drawing.Color.LightGray
            End If

            Dim tRow As New TableRow
            Dim tCounterCell As New TableCell
            Dim tPlayerNameCell As New TableCell

            Dim tCounterLabel As New Label
            Dim tPlayerDDL As New DropDownList

            tCounterLabel.Text = playerCount.ToString + "."
            tCounterCell.CssClass = "cfont10"

            tPlayerDDL = loadUnassignedPlayersDDL()
            tPlayerDDL.ID = "PERSON_ID" + playerCount.ToString
            tPlayerDDL.SelectedIndex = 0
            tPlayerDDL.Width = 200

            tCounterCell.Controls.Add(tCounterLabel)
            tCounterCell.HorizontalAlign = HorizontalAlign.Right
            tPlayerNameCell.Controls.Add(tPlayerDDL)
            tPlayerNameCell.HorizontalAlign = HorizontalAlign.Left

            tRow.Cells.Add(tCounterCell)
            tRow.Cells.Add(tPlayerNameCell)
            tRow.BackColor = BGCOLOR
            Me.teamTable.Rows.Add(tRow)

        Next

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        'lets loop through and make the adjustments to this team
        'first clear current team except for the captain

        'reset
        Me.successLabel.Text = ""

        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        'Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        SQL = "DELETE FROM TEAM_MEMBER_TBL " _
            & "WHERE TEAM_ID='" & Me.teamID.ToString & "' " _
            & "AND PERSON_ID <> '" & Me.personID.ToString & "'"
        rs = New SqlCommand(SQL, cn)
        'Response.Write(SQL & "<br />")
        rs.ExecuteNonQuery()

        For counter As Integer = 1 To 10
            Dim PID As String
            PID = Request.Form("PERSON_ID" + counter.ToString)

            'dont insert blank rows or the disabled row of the captain
            'and dont insert person twice if the captain picked them twice
            'from the DDL
            If PID <> "" And PID <> "0" Then
                SQL = "INSERT INTO TEAM_MEMBER_TBL(TEAM_ID,PERSON_ID,CAPTAIN_IND,CERTIFIED_REF_IND) " _
                    & "SELECT '" & Me.teamID.ToString & "', " _
                    & "'" & PID & "', " _
                    & "'N', " _
                    & "'N' " _
                    & "WHERE '" & PID & "' NOT IN " _
                    & "(SELECT PERSON_ID FROM TEAM_MEMBER_TBL " _
                     & "WHERE TEAM_ID='" & Me.teamID.ToString & "' " _
                     & "AND PERSON_ID='" & PID & "')"
                'Response.Write(SQL & "<br />")
                rs = New SqlCommand(SQL, cn)
                rs.ExecuteNonQuery()
            End If
        Next
        cn.Close()

        'Response.End()
        Me.successLabel.Text = "You have successfully updated your team."
        Me.successLabel.ForeColor = Drawing.Color.Green
        getTeamRoster()

    End Sub


End Class
