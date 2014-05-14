Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class MyTeams
    Inherits System.Web.UI.Page

    Public personID As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'TESTING!
        'Session("PERSON_ID") = 511
        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString()
        Else
            personID = 0
        End If

        If personID = 0 Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If

        If Not IsPostBack Then

            If validateRegistration() Then
                getMyTeams()
                getEventCodeDropDownList()
                getMessageLabel()
            Else
                Me.messageLabel.Text = "You have not registered for an event yet. Please click <a href=""MyEvents.aspx"">here</a> to register."
                Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
                teamPanel.Visible = False
                Me.createTeamPanel.Visible = False
            End If

        End If

    End Sub

    Protected Function validateRegistration() As Boolean
        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        SQL = "SELECT EVENT_CD FROM REGISTRATION_TBL " _
            & "WHERE PERSON_ID='" & personID.ToString & "' " _
            & "AND EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y') " _
            & "AND REGISTRATION_IND = 'Y'"

        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader

        If rsreply.HasRows Then
            validateRegistration = True
        Else
            validateRegistration = False
        End If

        rsreply.Close()
        cn.Close()

    End Function

    Protected Sub getMessageLabel()
        'If Me.teamPanel.Visible = False And Me.createTeamPanel.Visible = False And validateRegistration() = True Then
        'Me.messageLabel.Text = "You have already been selected on team(s) for your registered event(s)."
        'ElseIf
        If validateRegistration() = False Then
            Me.messageLabel.Text = "You have not registered for an event yet. Please click <a href=""MyEvents.aspx"">here</a> to register."
        ElseIf Me.teamPanel.Visible Then
            Me.messageLabel.Text = "If you are a team captain, click on a team name to manage your team roster."
        Else
            Me.messageLabel.Text = ""
        End If
    End Sub

    Protected Sub getMyTeams()
        'start with cleared table
        For counter As Integer = 2 To Me.teamTable.Rows.Count - 1
            Me.teamTable.Rows(counter).Controls.Clear()
            'Me.teamTable.Rows(counter).Dispose
        Next

        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        If InStr(Request.ServerVariables("HTTP_REFERER"), "MyCGVA_DEV", CompareMethod.Text) > 0 Then
            SQL = "MYCGVA_MYTEAMS_DEV @PERSON_ID='" & personID.ToString & "'"
        Else
            SQL = "MYCGVA_MYTEAMS @PERSON_ID='" & personID.ToString & "'"
        End If
        'Response.Write(SQL)
        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader

        If rsreply.HasRows Then
            Me.teamPanel.Visible = True

            While rsreply.Read
                Dim tRow As New TableRow
                Dim tEventCell As New TableCell
                Dim tTeamNameCell As New TableCell
                Dim tCaptainNameCell As New TableCell
                Dim tCaptainEmailCell As New TableCell

                Dim tEventLabel As New Label
                Dim tTeamNameHyperlink As New HyperLink
                Dim tTeamNameLabel As New Label
                Dim tCaptainNameLabel As New Label
                Dim tCaptainEmailLabel As New Label

                tEventLabel.Text = rsreply(1)
                tEventCell.Controls.Add(tEventLabel)
                tEventCell.HorizontalAlign = HorizontalAlign.Center
                tEventCell.CssClass = "cfont10"

                tTeamNameHyperlink.NavigateUrl = "TeamMaintenance.aspx?ID=" & rsreply(0).ToString & "&EVENT_CD=" & rsreply(3)
                tTeamNameHyperlink.Text = rsreply(2)
                tTeamNameLabel.Text = rsreply(2)

                'captains get the hyperlink to edit the team
                If rsreply(4) = "Y" Then
                    tTeamNameCell.Controls.Add(tTeamNameHyperlink)
                    Me.submitButton.Visible = True
                Else
                    tTeamNameCell.Controls.Add(tTeamNameLabel)
                    Me.submitButton.Visible = False
                End If

                tTeamNameCell.HorizontalAlign = HorizontalAlign.Center
                tTeamNameCell.CssClass = "cfont10"

                tCaptainNameLabel.Text = rsreply(5)
                tCaptainNameCell.Controls.Add(tCaptainNameLabel)
                tCaptainNameCell.HorizontalAlign = HorizontalAlign.Center
                tCaptainNameCell.CssClass = "cfont10"

                tCaptainEmailLabel.Text = rsreply(6)
                tCaptainEmailCell.Controls.Add(tCaptainEmailLabel)
                tCaptainEmailCell.HorizontalAlign = HorizontalAlign.Center
                tCaptainEmailCell.CssClass = "cfont10"

                tRow.Cells.Add(tEventCell)
                tRow.Cells.Add(tTeamNameCell)
                tRow.Cells.Add(tCaptainNameCell)
                tRow.Cells.Add(tCaptainEmailCell)
                Me.teamTable.Rows.Add(tRow)

            End While

        Else
            'add headers, with no team names assigned
            Dim tRow As New TableRow
            Dim tEventCell As New TableCell

            Dim tEventLabel As New Label

            tEventLabel.Text = "You have not been added to any teams at this time."
            tEventCell.Controls.Add(tEventLabel)
            tEventCell.HorizontalAlign = HorizontalAlign.Center
            tEventCell.CssClass = "cfont10"
            tEventCell.ColumnSpan = 4
            tEventCell.ForeColor = System.Drawing.Color.Firebrick

            tRow.Cells.Add(tEventCell)
            Me.teamTable.Rows.Add(tRow)
            'Me.teamPanel.Visible = False
        End If
        cn.Close()
    End Sub

    Protected Sub getEventCodeDropDownList()
        'NOT NEEDED
        'Dim vS2 As String = ""
        'Dim vS4 As String = ""
        'Dim vS6 As String = ""

        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        SQL = "SELECT r.EVENT_CD,e.EVENT_SHORT_DESC FROM REGISTRATION_TBL r " _
            & "LEFT JOIN EVENT_TBL e on r.EVENT_CD = e.EVENT_CD " _
            & "WHERE PERSON_ID='" & personID.ToString & "' " _
            & "AND r.EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y') " _
            & "AND REGISTRATION_IND = 'Y' " _
            & "AND r.EVENT_CD NOT IN " _
            & "(SELECT dt.EVENT_CD FROM DIVISION_TBL dt " _
            & "LEFT JOIN TEAM_TBL tt on dt.DIVISION_ID = tt.DIVISION_ID " _
            & "LEFT JOIN TEAM_MEMBER_TBL tmt on tt.TEAM_ID = tmt.TEAM_ID " _
            & "WHERE tmt.PERSON_ID='" & personID.ToString & "')"
        ' AND tmt.CAPTAIN_IND='Y' dont use this for the DDL, 
        'dont want people to create teams if they are already on a team for this event

        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader

        'start with empty DDL
        Me.eventCodeDropDownList.Items.Clear()
        'add base item
        Me.eventCodeDropDownList.Items.Add(New ListItem("SELECT", ""))

        If rsreply.HasRows Then
            Me.createTeamPanel.Visible = True
            While rsreply.Read
                Me.eventCodeDropDownList.Items.Add(New ListItem(rsreply("EVENT_SHORT_DESC"), rsreply("EVENT_CD")))
            End While

        Else
            Me.createTeamPanel.Visible = False
        End If

        rsreply.Close()
        cn.Close()

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        'Response.Write(Me.eventCodeDropDownList.SelectedValue)
        'Response.Write(Me.eventCodeDropDownList.SelectedItem.ToString)
        'Response.Write(Me.eventCodeDropDownList.SelectedIndex.ToString)

        Me.messageLabel.Text = ""
        Me.messageLabel.ForeColor = Drawing.Color.Black

        If Me.eventCodeDropDownList.SelectedValue = "" Then
            Me.messageLabel.Text = "Please select an event for the team you are creating."
            Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
            Exit Sub
        ElseIf Trim(Me.teamNameTextBox.Text) = "" Then
            Me.messageLabel.Text = "Please enter a name for the team you are creating."
            Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
            Exit Sub
        End If

        'ok. we have a valid event and a possible team name
        'needs to be validated against teams already created
        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader
        Dim eventCode As String = Me.eventCodeDropDownList.SelectedValue
        Dim tempTeamName As String = Replace(Trim(Me.teamNameTextBox.Text), "'", "''")

        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        SQL = "SELECT COUNT(*) FROM TEAM_TBL tt " _
            & "LEFT JOIN DIVISION_TBL dt on tt.DIVISION_ID = dt.DIVISION_ID " _
            & "WHERE dt.EVENT_CD = '" & eventCode & "' " _
            & "AND UPPER(tt.TEAM_NAME) = '" & UCase(tempTeamName) & "'"
        'Response.Write(SQL)
        'Response.End()
        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader
        rsreply.Read()

        'team already exists?
        If rsreply(0) > 0 Then
            rsreply.Close()
            cn.Close()
            Me.messageLabel.Text = "The team name you entered already exists for this event. Please enter a different name."
            Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
            Exit Sub
        Else
            'team name is unique, let's add it to this event after finding the temp division ID
            rsreply.Close()
            Dim tempDivisionID As Integer = 0
            Dim tempTeamNumber As Integer = 0
            Dim tempTeamNumberStr As String = ""
            Dim tempTeamID As Integer = 0

            SQL = "SELECT TOP 1 DIVISION_ID FROM DIVISION_TBL " _
                & "WHERE EVENT_CD='" & eventCode & "'"
            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader


            If rsreply.Read Then
                tempDivisionID = rsreply(0)
                rsreply.Close()
            Else
                'ERROR: no divisions have been set up for this event yet
                rsreply.Close()
                cn.Close()
                Me.messageLabel.Text = "We are sorry, there are no divisions set up for this event at this time. Please contact the <a href=""mailto:support@cgva.org"">CGVA webmaster</a> about this issue."
                Me.messageLabel.ForeColor = System.Drawing.Color.Firebrick
                Exit Sub
            End If

            'get last team created for this event/division 
            'for incrementation of the team code
            SQL = "SELECT COUNT(*) FROM TEAM_TBL " _
                & "WHERE DIVISION_ID = '" & tempDivisionID.ToString & "'"
            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            rsreply.Read()
            tempTeamNumber = rsreply(0) + 1

            'for website display reasons
            If tempTeamNumber < 10 Then
                tempTeamNumberStr = "T0" + tempTeamNumber.ToString
            Else
                tempTeamNumberStr = "T" + tempTeamNumber.ToString
            End If
            rsreply.Close()

            'enter new team
            SQL = "INSERT INTO TEAM_TBL(DIVISION_ID,TEAM_CD,TEAM_NAME) " _
                & "SELECT '" & tempDivisionID.ToString & "', " _
                & "'" & tempTeamNumberStr & "', " _
                & "'" & tempTeamName & "' " _
                & "SELECT SCOPE_IDENTITY()"
            'Response.Write(SQL)
            'Response.End()
            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            rsreply.Read()
            tempTeamID = rsreply(0)
            rsreply.Close()

            'add this person as captain to this team in the TEAM_MEMBER_TBL
            SQL = "INSERT INTO TEAM_MEMBER_TBL(PERSON_ID,TEAM_ID,CAPTAIN_IND,CERTIFIED_REF_IND) " _
                & "SELECT '" & personID.ToString & "', " _
                & "'" & tempTeamID.ToString & "', " _
                & "'Y', " _
                & "'N'"
            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()
            cn.Close()

            'reload the event DDL if necessary
            getMyTeams()
            getEventCodeDropDownList()

            getMessageLabel()


        End If

    End Sub

End Class
