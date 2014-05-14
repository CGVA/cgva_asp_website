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

Public Class PayPal
    Inherits System.Web.UI.Page


    Protected cmd As String = "_xclick"
    Protected business As String = ""
    Protected item_name As String = AppSettings("item_name")
    Protected item_number As String = ""    'used to hold person ID of the transaction

    Protected amount As String
    'Protected return_url As String = AppSettings("ReturnUrl")
    'JPC TEST'
    Protected return_url As String = "http://cgva.org/MyCGVA/MyEvents6.aspx"
    Protected notify_url As String = AppSettings("NotifyUrl")
    'Protected notify_url As String = ""
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load


        ' determining the URL to work with depending on whether sandbox or a real PayPal account should be used
        If AppSettings("UseSandbox").ToString = "true" Then
            URL = "https://www.sandbox.paypal.com/cgi-bin/webscr"
            business = AppSettings("devBusinessEmail")
            paypaluser = AppSettings("devUSER")
            pwd = AppSettings("devPassword")
            signature = AppSettings("devSIGNATURE")

        Else
            URL = "https://www.paypal.com/cgi-bin/webscr"
            business = AppSettings("prodBusinessEmail")
            paypaluser = AppSettings("prodUSER")
            pwd = AppSettings("prodPassword")
            signature = AppSettings("prodSIGNATURE")
        End If

        'This parameter determines the was information about successful transaction 
        'will be passed to the script
        ' specified in the return_url parameter.
        ' "1" - no parameters will be passed.
        ' "2" - the POST method will be used.
        ' "0" - the GET method will be used. 
        ' The parameter is "0" by deault.
        If AppSettings("SendToReturnURL").ToString = "true" Then
            rm = "2"
        Else
            rm = "1"
        End If

        ' the total cost of the cart
        amount = "150"
        custom = "CUSTOM"

        'for Kevins reconciliation
        item_number = "TEAMNAME"

    End Sub
End Class

