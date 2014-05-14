Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class MyEvents
    Inherits System.Web.UI.Page

    Public personID As Integer = 0
    Private boolEventsAvailForReg As Boolean = False
    Private intCheckBoxCheckedCount As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString()
        Else
            personID = 0
        End If

        If personID = 0 Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If

        'If Not IsPostBack Then


        Dim SQL As String
        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")
        Dim tempPassword As String = ""

        'Dim vLeague As String = ""
        Dim vS2 As String = ""
        Dim vS4 As String = ""
        Dim vS6 As String = ""
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsreply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        SQL = "if (select count(*) from EVENT_TBL e " _
           & "LEFT JOIN REGISTRATION_TBL r ON e.EVENT_CD = r.EVENT_CD " _
           & "WHERE e.EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y') " _
           & "AND r.PERSON_ID='" & Session("PERSON_ID") & "' " _
           & "AND REGISTRATION_IND='Y') > 0 " _
           & "BEGIN " _
               & "SELECT e.EVENT_SHORT_DESC, " _
               & "e.EVENT_CD,'Y' as 'REGISTRATION_IND' " _
               & "FROM EVENT_TBL e " _
               & "WHERE e.EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y') " _
               & "AND e.EVENT_CD IN " _
                    & "(SELECT EVENT_CD " _
                    & "FROM REGISTRATION_TBL " _
                    & "WHERE PERSON_ID = '" & Session("PERSON_ID") & "' " _
                    & "AND REGISTRATION_IND='Y') " _
               & "ORDER BY EVENT_SHORT_DESC " _
           & "END " _
        & "ELSE " _
        & "BEGIN " _
               & "SELECT DISTINCT(e.EVENT_SHORT_DESC), " _
               & "e.EVENT_CD,'N' as 'REGISTRATION_IND' " _
               & "FROM EVENT_TBL e " _
               & "LEFT JOIN REGISTRATION_TBL r ON e.EVENT_CD = r.EVENT_CD " _
               & "WHERE e.EVENT_CD IN (SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y') " _
        & "End"
        '& "AND r.PERSON_ID='" & Session("PERSON_ID") & "' " _
        '& "e.EVENT_CD,IsNull(r.REGISTRATION_IND,'N') as 'REGISTRATION_IND' " _

        rs = New SqlCommand(SQL, cn)
        rsreply = rs.ExecuteReader

        If rsreply.HasRows Then

            'loop through open registration events
            'determine if this person has registered for each event or not
            'display checkbox if they can register, label if they already have
            While rsreply.Read
                Dim tRow As New TableRow()
                Me.eventTable.Rows.Add(tRow)
                Dim tCell As New TableCell()
                tRow.Cells.Add(tCell)

                If rsreply("REGISTRATION_IND") = "Y" Then
                    Dim vLabel As New Label
                    vLabel.Text = rsreply("EVENT_SHORT_DESC") + " - REGISTRATION COMPLETE"
                    vLabel.Font.Bold = True
                    vLabel.ForeColor = Drawing.Color.Green
                    vLabel.CssClass = "cfont10"
                    tCell.Controls.Add(vLabel)
                Else
                    boolEventsAvailForReg = True
                    Dim vCheckbox As New RadioButton
                    'Dim vCheckbox As New CheckBox
                    vCheckbox.ID = rsreply("EVENT_CD")
                    vCheckbox.CssClass = "cfont10"
                    vCheckbox.Text = rsreply("EVENT_SHORT_DESC")
                    tCell.Controls.Add(vCheckbox)
                    AddHandler vCheckbox.CheckedChanged, AddressOf checkboxCheckedCount
                End If

            End While

        Else
            'no events are open for registration
            Dim tRow As New TableRow()
            Me.eventTable.Rows.Add(tRow)
            Dim tCell As New TableCell()
            tRow.Cells.Add(tCell)

            Dim vLabel As New Label
            vLabel.Text = "NO EVENTS ARE OPEN FOR REGISTRATION AT THIS TIME."
            vLabel.Font.Bold = True
            vLabel.ForeColor = Drawing.Color.Red
            vLabel.CssClass = "cfont10"
            tCell.Controls.Add(vLabel)

        End If
        cn.Close()

        If boolEventsAvailForReg Then
            'add submit button
            Dim tRow As New TableRow()
            Me.eventTable.Rows.Add(tRow)
            Dim tCell As New TableCell()
            tRow.Cells.Add(tCell)
            Dim vButton As New Button

            vButton.Text = "Continue"
            vButton.ID = "submitButton"
            AddHandler vButton.Click, AddressOf submitButton_Click
            tCell.Controls.Add(vButton)

        End If

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If intCheckBoxCheckedCount = 0 Then
            Me.messageLabel.Text = "Please select at least one event to continue with registration."
        Else
            'Context.Items.Add("League", Me.League.Checked)
            'Context.Items.Add("SB", Me.SB.Checked)
            Server.Transfer("MyEvents2.aspx", True)
        End If
    End Sub

    Protected Sub checkboxCheckedCount(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.checked Then
            intCheckBoxCheckedCount += 1
        Else
            intCheckBoxCheckedCount -= 1
        End If
    End Sub
End Class
