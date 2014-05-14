Imports System.Data.SqlClient

Partial Class VerifyPassword
    Inherits System.Web.UI.Page

    Public personID As String = "0"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session("PERSON_ID") = ""

        If Not (Session("PERSON_ID") Is Nothing) Then
            personID = Session("PERSON_ID").ToString
            Me.USERNAME.Text = Session("USERNAME")
        Else
            personID = "0"
        End If

        If personID = "0" Then
            Response.Redirect("Login.aspx?timeout=true")
            Exit Sub
        End If

        'Response.Write("UN:" & Session("USERNAME"))

    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click

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
            & "ANSWER = '" & securityAnswer & "', " _
            & "TEMP_PASSWORD = '' " _
            & "WHERE PERSON_ID = '" & Session("PERSON_ID") & "'"

            rs = New SqlCommand(SQL, cn)
            rs.ExecuteNonQuery()
            cn.Close()

            'reset session variable
            Session("TEMP_PASSWORD") = ""
        End If

        Me.messageLabel.Text = "<font class='cfontSuccess10'>Congratulations, you will be able to access all areas of MyCGVA. Click <a href='MyProfile.aspx'>here</a> to verify your complete profile information, or select an option from the left hand navigation bar.</font>"
        Me.submitButton.Visible = False

    End Sub

    Sub closeRSCNConnection(ByVal rsreply As SqlDataReader, ByVal cn As SqlConnection)
        rsreply.Close()
        cn.Close()
    End Sub

End Class
