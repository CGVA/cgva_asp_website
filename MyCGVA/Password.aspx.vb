Imports System.Data.SqlClient
Imports System.Net.Mail

Imports System.Net
Imports System.IO
Imports System.Globalization
Imports System.Configuration.ConfigurationManager
Imports System.Data
Imports System.Xml


Partial Class Password
    Inherits System.Web.UI.Page

    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")
    Public personID As Integer = 0

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

        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsReply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()

        If Not IsPostBack Then

            SQL = "SELECT USERNAME,PASSWORD,IsNull(SECURITY_QUESTION,0) as 'SECURITY_QUESTION',ANSWER " _
                & "FROM USER_LOGIN_TBL " _
                & "WHERE PERSON_ID = '" & Session("PERSON_ID") & "'"
            rs = New SqlCommand(SQL, cn)
            rsReply = rs.ExecuteReader
            If rsReply.Read Then
                Me.USERNAME.Text = rsReply("USERNAME")
                Me.PASSWORD1.Text = rsReply("PASSWORD")
                Me.SECURITY_QUESTION.SelectedValue = rsReply("SECURITY_QUESTION")
                Me.SECURITY_ANSWER.Text = rsReply("ANSWER")
            End If
            'PWGenerator.PasswordGenerator.CreatePasswords()
        End If

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        'update password information for this user
        Dim strPassword1 = Trim(Me.PASSWORD1.Text)
        Dim strPassword2 = Trim(Me.PASSWORD2.Text)
        Dim securityQuestion = Me.SECURITY_QUESTION.SelectedValue
        Dim securityQuestionText = Me.SECURITY_QUESTION.SelectedItem.Text
        Dim securityAnswer = Trim(Me.SECURITY_ANSWER.Text)

        If UCase(strPassword1) = UCase(Me.USERNAME.Text) Then
            Me.messageLabel.Text = "Your password cannot be the same as your user name."
            Exit Sub
        ElseIf strPassword1 = "" Then
            Me.messageLabel.Text = "Please enter a password."
            Exit Sub
        ElseIf strPassword1.Length < 6 Then
            Me.messageLabel.Text = "Please enter a password between 6 and 20 characters in length."
            Exit Sub
        ElseIf strPassword2 = "" Then
            Me.messageLabel.Text = "Please verify the password."
            Exit Sub
        ElseIf strPassword1 <> strPassword2 Then
            Me.messageLabel.Text = "Please make sure the password and the re-typed password are identical."
            Exit Sub
        ElseIf securityQuestion = 0 Then
            Me.messageLabel.Text = "Please select a security question."
            Exit Sub
        ElseIf securityAnswer = "" Then
            Me.messageLabel.Text = "Please supply an answer to the selected security question."
            Exit Sub
        Else

            Dim cn As SqlConnection
            Dim rs As SqlCommand
            'Dim rsReply As SqlDataReader
            cn = New SqlConnection
            cn.ConnectionString = connStr
            cn.Open()

            'update this user's password information and remove the temp password stored in the database
            SQL = "UPDATE USER_LOGIN_TBL " _
            & "SET PASSWORD = '" & strPassword1 & "', " _
            & "SECURITY_QUESTION = '" & securityQuestion.ToString & "', " _
            & "ANSWER = '" & securityAnswer & "' " _
            & "WHERE PERSON_ID = '" & Session("PERSON_ID") & "'"

            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()
            cn.Close()

            'reset session variable
            Session("TEMP_PASSWORD") = ""
        End If

        Me.messageLabel.Text = "<font class='cfontSuccess10'>Your password information has been updated successfully.</font>"
        'Me.submitButton.Visible = False
    End Sub

End Class
