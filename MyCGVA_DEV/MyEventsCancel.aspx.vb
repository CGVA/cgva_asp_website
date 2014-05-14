
Partial Class MyEventsCancel
    Inherits System.Web.UI.Page

    Dim personID As String = "0"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            Dim aryTextFile() As String
            aryTextFile = Request.QueryString("custom").Split("~")

            personID = aryTextFile(0)
            Session("PERSON_ID") = personID

        Catch ex As Exception
            'Response.Redirect("Login.aspx")
        End Try
        messageLabel.Text = "Your registration request has been canceled. Click <a href='http://www.cgva.org'>here</a> to return to the main CGVA site."

    End Sub
End Class
