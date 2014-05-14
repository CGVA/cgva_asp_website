Imports System.Data.SqlClient

Partial Class RegistrationCancel
    Inherits System.Web.UI.Page

    'delete the transaction from the team table
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        '        Dim ID As Integer = 0

        '        Try
        '        ID = Request.QueryString("ID")
        '        Catch ex As Exception

        '        End Try
        '        Dim SQL As String = ""
        '        Dim connStr As String = ConfigurationManager.AppSettings("ConnectionString")

        '        Dim cn As SqlConnection
        '        Dim rs As SqlCommand
        '        cn = New SqlConnection
        '        cn.ConnectionString = connStr
        '        cn.Open()

        '        SQL = "DELETE FROM LASTDIG_TEAM_TBL " _
        '                & "WHERE ID='" & ID.ToString & "'"

        '        rs = New SqlCommand(SQL, cn)
        '        rs.ExecuteNonQuery()

        '        cn.Close()
        '        Response.Clear()

        '        If Not IsPostBack Then
        'Response.Redirect("http://lastdigindenver.org/Registration.aspx?success=false", True)
        'Response.Status = "301 Moved Permanently";
        'Response.AddHeader("", "http://lastdigindenver.org/Registration.aspx?success=false")

        '        End If

    End Sub

    Private Sub unloadPage(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        'Response.Write("HERE")
        'Response.End()
        'Response.Redirect("http://lastdigindenver.org/Registration.aspx?success=false")
        'Server.Execute("Registration.aspx?success=false", False)
    End Sub

End Class
