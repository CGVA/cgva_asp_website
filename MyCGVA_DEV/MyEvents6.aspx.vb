Imports System.Data.SqlClient
Imports System.Net.Mail

Partial Class MyEvents6
    Inherits System.Web.UI.Page

    Private vVerifyInfo As String
    Private vVerifyWaiverInfo As String
    Private vLeague As String
    Private vSB As String

    Private vCAPTAIN_IND As String
    Private vCAPTAIN_INTEREST As String
    Private vCROSS_DIVISION_IND As String

    Private vQ1 As String
    Private vQ2 As String
    Private vQ3 As String
    Private vQ4 As String
    Private vQ5 As String

    Private vrequest_id As String

    Public Const EMAIL_SERVER As String = "relay-hosting.secureserver.net"
    Public personID As String = "0"
    Private SQL As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)

        messageLabel.Text = "Congratulations! You have succussfully registered for the CGVA league(s)."
        messageLabel.Text += "<p />Click <a href='http://www.cgva.org'>here</a> to return to the main CGVA site."

    End Sub
End Class
