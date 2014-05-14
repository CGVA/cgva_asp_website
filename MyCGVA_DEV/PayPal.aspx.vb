'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  File:           PayPal.aspx.vb
'
'  Facility:       The unit contains the PayPal class
'
'  Abstract:       This class is intended for interacting with PayPal with the
'                  help of the form of the payment request this class creates.
'
'  Environment:    VC 8.0
'
'  Author:         KB_Soft Group Ltd.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Imports System.Configuration.ConfigurationManager
Imports System.Data.SqlClient

Partial Class PayPal
    Inherits System.Web.UI.Page

    Protected cmd As String = "_xclick"
    Protected business As String = ""
    'Protected item_name As String = AppSettings("item_name")
    Protected item_name As String = "2013 - Volleypalooza"
    Protected item_number As String = ""    'used to hold person ID of the transaction

    Protected amount As String
    'Protected return_url As String = AppSettings("ReturnUrl")
    'JPC TEST'
    Protected return_url As String = "http://cgva.org/MyCGVA/MyEvents6.aspx"
    'Protected notify_url As String = AppSettings("NotifyUrl")
    'JPC TEST'
    Protected notify_url As String = "http://cgva.org/MyCGVA/IPNHandler.aspx"
    'JPC TEST'
    Protected cancel_url As String = "http://cgva.org/MyCGVA/MyEventsCancel.aspx"
    Protected currency_code As String = AppSettings("CurrencyCode")
    Protected paypaluser As String = ""
    Protected pwd As String = ""
    Protected signature As String = ""
    Protected no_shipping As String = "1"
    Protected URL As String
    Protected custom As String = ""
    Protected invoice As String = ""
    Protected rm As String
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")
    Private eventArray As New ArrayList
    Private SQL As String


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load


        ' determining the URL to work with depending on whether sandbox or a real PayPal account should be used
        'If AppSettings("UseSandbox").ToString = "true" Then
        If InStr(Request.ServerVariables("HTTP_REFERER"), "MyCGVA_DEV", CompareMethod.Text) > 0 Then
            URL = "https://www.sandbox.paypal.com/cgi-bin/webscr"
            business = AppSettings("devBusinessEmail")
            paypaluser = AppSettings("devUSER")
            pwd = AppSettings("devPWD")
            signature = AppSettings("devSIGNATURE")
            return_url = "http://cgva.org/MyCGVA_DEV/MyEvents6.aspx"
            notify_url = "http://cgva.org/MyCGVA_DEV/IPNHandler.aspx"
            cancel_url = "http://cgva.org/MyCGVA_DEV/MyEventsCancel.aspx"
        Else
            URL = "https://www.paypal.com/cgi-bin/webscr"
            business = AppSettings("prodBusinessEmail")
            paypaluser = AppSettings("prodUSER")
            pwd = AppSettings("prodPWD")
            signature = AppSettings("prodSIGNATURE")
            return_url = "http://cgva.org/MyCGVA/MyEvents6.aspx"
            notify_url = "http://cgva.org/MyCGVA/IPNHandler.aspx"
            cancel_url = "http://cgva.org/MyCGVA/MyEventsCancel.aspx"
        End If

        'This parameter determines the was information about successful transaction 
        'will be passed to the script
        ' specified in the return_url parameter.
        ' "1" - no parameters will be passed.
        ' "2" - the POST method will be used.
        ' "0" - the GET method will be used. 
        ' The parameter is "0" by deault.

        'If AppSettings("SendToReturnURL").ToString = "true" Then
        rm = "2"
        'Else
        'rm = "1"
        'End If
        'rm = "0"

        ' the total cost of the cart
        amount = Session("Amount")
        ' the identifier of the payment request
        '08/03/2012 find out which event they registered for
        Dim cn As SqlConnection
        Dim rs As SqlCommand
        Dim rsReply As SqlDataReader
        cn = New SqlConnection
        cn.ConnectionString = connStr
        cn.Open()
        SQL = "SELECT EVENT_CD FROM EVENT_TBL " _
            & "WHERE EVENT_CD IN " _
            & "(SELECT EVENT_CD FROM EVENT_TBL WHERE OPEN_REGISTRATION_IND='Y')"

        rs = New SqlCommand(SQL, cn)
        rsReply = rs.ExecuteReader

        If rsReply.HasRows Then
            While rsReply.Read
                'Response.Write(rsReply("EVENT_CD"))
                'Response.End()
                Me.eventArray.Add(rsReply("EVENT_CD"))
            End While

            rsReply.Close()
        Else
            'shouldnt get here, but redirect back to MyEvents if we do
            cn.Close()
            Response.Redirect("MyEvents.aspx")
            Response.End()
        End If

        'create hidden fields for any chosen events, maintaining the values for next page
        Dim tempString As String = ""

        'hardcoded for summer leagues
        'Dim s4, s6 As Boolean
        's4 = False
        's6 = False

        '8/3/2012 for now, let's  stop exit for loop allowing for only one league registration
        For Each s As String In Me.eventArray

            'Response.Write(s & "<br />")

            Dim vHiddenField As New HiddenField
            vHiddenField.ID = s
            vHiddenField.Value = ""

            'Response.Write("count:" & Request.Form.Count.ToString & "<br />")

            For i = 0 To Request.Form.Count - 1
                'Response.Write(Request.Form.AllKeys(i) & ":")
                'Response.Write(Request.Form(i) & "<br />")
                tempString = Request.Form.AllKeys(i)

                If s = tempString And Request.Form(i) <> "" Then
                    'vHiddenField.Value = Request.Form(i)

                    custom = Session("PERSON_ID") + "~" _
                                    & s + "~" _
                                    & Session("Amount") + "~" _
                                    & Session("verifyInfo") + "~" _
                                    & Session("verifyWaiverInfo") + "~" _
                                    & Session("CAPTAIN_IND_League") + "~" _
                                    & Session("Q5")

                    Exit For
                End If

            Next


        Next


        'JPC 12/17/10 return_url += "?custom=" & custom

        'for treasurer reconciliation
        item_number = Session("NAME")

    End Sub
End Class

