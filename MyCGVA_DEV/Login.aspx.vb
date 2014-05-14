Imports System.Data.SqlClient

Partial Class Login
    Inherits System.Web.UI.Page

    Private USER_NAME As String
    Private PWORD As String
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session("PERSON_ID") = ""

        If Request.QueryString("timeout") = "true" Then
            Me.ErrorLabel.Text = "Sorry, your session timed out. Please log in again."
        End If

    End Sub

    Protected Sub Page_Unload(ByVal Sender As Object, ByVal E As EventArgs) Handles Me.Unload

        'Page.SetFocus(Me.USERNAME)
        'Me.USERNAME.Focus()

    End Sub


    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click

        USER_NAME = UCase(Replace(Me.USERNAME.Text, "'", "''"))
        PWORD = UCase(Replace(Me.PASSWORD.Text, "'", "''"))

        If Trim(USER_NAME) = "" Or Trim(PWORD) = "" Then
            ErrorLabel.Text = "Please enter your user name/password."
            Me.USERNAME.Focus()
        Else
            'Response.Write("HERE")
            'verify login info
            Dim tempPassword As String = ""

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            Dim rsreply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            SQL = "SELECT ul.PERSON_ID,ul.USERNAME,p.LAST_NAME,p.FIRST_NAME,IsNull(ul.TEMP_PASSWORD,'') as 'TEMP_PASSWORD' " _
            & "FROM USER_LOGIN_TBL ul " _
            & "LEFT JOIN db_accessadmin.[PERSON_TBL] p " _
            & "ON ul.PERSON_ID = p.PERSON_ID " _
            & "WHERE UPPER(USERNAME) = '" & UCase(USER_NAME) & "' " _
            & "AND (UPPER(PASSWORD) = '" & UCase(PWORD) & "' " _
            & "OR UPPER(TEMP_PASSWORD) = '" & UCase(PWORD) & "')"

            rs = New SqlCommand(SQL, cn)
            rsreply = rs.ExecuteReader
            'set rs = cn.(SQL)

            If rsreply.Read Then
                'information found = proceed
                Session("PERSON_ID") = rsreply("PERSON_ID").ToString
                Session("NAME") = rsreply("FIRST_NAME") & " " & rsreply("LAST_NAME")
                'Session("FIRST_NAME") = rsreply("FIRST_NAME")
                'Session("LAST_NAME") = rsreply("LAST_NAME")
                Session("USERNAME") = rsreply("USERNAME")
                tempPassword = rsreply("TEMP_PASSWORD")
                Session("TEMP_PASSWORD") = tempPassword

                'record login attempt
                recordLogin("Y")
                closeRSCNConnection(rsreply, cn)
                'Response.Redirect("MyProfile.aspx")

                'check for temp password (1st time accessing the site)
                If tempPassword = "" Then
                    Response.Redirect("MyProfile.aspx")
                Else
                    Response.Redirect("VerifyPassword.aspx")
                End If
            Else
                closeRSCNConnection(rsreply, cn)
                'incorrect login info - redirect with message
                ErrorLabel.Text = "Sorry, the user name/password entered were not found. Please try again."
                Me.USERNAME.Focus()

                'record login attempt
                recordLogin("N")
            End If

        End If


    End Sub

    Sub closeRSCNConnection(ByVal rsreply As SqlDataReader, ByVal cn As SqlConnection)
        rsreply.Close()
        cn.Close()
    End Sub

    Sub recordLogin(ByVal success As Char)
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        SQL = "INSERT INTO USER_LOGIN_LOG_ENTRY_TBL(" _
        & "ENTRY_USERNAME, " _
        & "ENTRY_PASSWORD, " _
        & "ENTRY_DATETIME, " _
        & "ENTRY_SUCCESS_IND) " _
        & "SELECT '" & USER_NAME & "', " _
        & "'" & PWORD & "', " _
        & "getdate(), " _
        & "'" & success & "'"

        'Response.Write(SQL & "<br />")
        'JPC - come back to this - double entries occurring for bad login info
        rs = New SqlCommand(SQL, cn)
        rs.ExecuteNonQuery()
        cn.Close()

    End Sub

End Class
