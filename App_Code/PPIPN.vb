
'Copyright 2008 - XDev Software, LLC 
'Use this code at your own risk. XDev Software is not responsible 
'for any damage that may come from using this code. This code is provided as is. 


' This code does not account for some of the variables that are passed back from IPN. 
' * Howevever, it does account for most of them. This code does not also account for all shopping 
' * cart functionality as you will have to handle the multiple items yourself. In addition, 
' * it does not handle refunds in anyway for IPN. 
' * Thanks again for using the code and any changes you feel that are needed please email to 
' * 
' * support@xdevsoftware.com 
' * 
' 


Imports System.Net
Imports System.Net.Mail
Imports System.IO
Imports System.Text
Imports System.Web
Imports System.Data.SqlClient

Public Class PPIPN
    Private connStr As String = ConfigurationManager.AppSettings("ConnectionString")
    Private SQL As String

    Private _txnID As String, _txnType As String, _paymentStatus As String, _receiverEmail As String, _itemName As String, _itemNumber As String, _
    _quantity As String, _invoice As String, _custom As String, _paymentGross As String, _payerEmail As String, _pendingReason As String, _
    _paymentDate As String, _paymentFee As String, _firstName As String, _lastName As String, _address As String, _city As String, _
    _state As String, _zip As String, _country As String, _countryCode As String, _addressStatus As String, _payerStatus As String, _
    _payerID As String, _paymentType As String, _notifyVersion As String, _verifySign As String, _response As String, _payerPhone As String, _
    _payerBusinessName As String, _business As String, _receiverID As String, _memo As String, _tax As String, _qtyCartItems As String, _
    _shippingMethod As String, _shipping As String

    Private _postUrl As String = ""
    Private _strRequest As String = ""
    Private _smtpHost As String, _fromEmail As String, _toEmail As String, _fromEmailPassword As String, _smtpPort As String

    ''' <summary> 
    ''' valid strings are "TEST" for sandbox use 
    ''' "LIVE" for production use 
    ''' "ELITE" for test use off of PayPal...avoid having to be logged into PayPal Developer 
    ''' </summary> 
    ''' <param name="mode"></param> 
    Public Sub New(ByVal mode As String)
        If mode.ToLower() = "test" Then
            Me.PostUrl = "https://www.sandbox.paypal.com/cgi-bin/webscr"
        ElseIf mode.ToLower() = "live" Then
            Me.PostUrl = "https://www.paypal.com/cgi-bin/webscr"
        ElseIf mode.ToLower() = "elite" Then
            Me.PostUrl = "http://www.eliteweaver.co.uk/testing/ipntest.php"
        Else
            Me.PostUrl = ""
        End If

        Me.fillProperties()
    End Sub

#Region "Properties"

    Private Property PostUrl() As String
        Get
            Return _postUrl
        End Get
        Set(ByVal value As String)
            _postUrl = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the reponse back from the http post back to PayPal. 
    ''' Possible values are "VERIFIED" or "INVALID" 
    ''' </summary> 
    Private Property Response() As String
        Get
            Return _response
        End Get
        Set(ByVal value As String)
            _response = value
        End Set
    End Property

    Private Property RequestLength() As String
        Get
            Return _strRequest
        End Get
        Set(ByVal value As String)
            _strRequest = value
        End Set
    End Property

    ''' <summary> 
    ''' Provide your outgoing email server to use are your SMTP host 
    ''' </summary> 
    Public Property SmtpHost() As String
        Get
            Return _smtpHost
        End Get
        Set(ByVal value As String)
            _smtpHost = value
        End Set
    End Property

    ''' <summary> 
    ''' Provide the port your outgoing SMTP host uses 
    ''' </summary> 
    Public Property SmtpPort() As String
        Get
            Return _smtpPort
        End Get
        Set(ByVal value As String)
            _smtpPort = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the email address that will show to the customer and you. This most likely 
    ''' needs to be a valid email address that your SMTP server will accept 
    ''' Examples would be something like no-reply@yourdomain.com 
    ''' </summary> 
    Public Property FromEmail() As String
        Get
            Return _fromEmail
        End Get
        Set(ByVal value As String)
            _fromEmail = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the password that the FromEmail property will use. This needs to be the password 
    ''' for the email account itself 
    ''' </summary> 
    Public Property FromEmailPassword() As String
        Get
            Return _fromEmailPassword
        End Get
        Set(ByVal value As String)
            _fromEmailPassword = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the email address that you use for yourself. This should be set to 
    ''' the email that is registered for your PayPal account. 
    ''' </summary> 
    Public Property ToEmail() As String
        Get
            Return _toEmail
        End Get
        Set(ByVal value As String)
            _toEmail = value
        End Set
    End Property

    ''' <summary> 
    ''' Email address or Account ID of the payment recipient. This is equivalent 
    ''' to the value of receiver_email if the payment is sent to the primary account 
    ''' , which is most cases it is. This value is that value of what is set in the button html 
    ''' markup. This value also get normalized to lowercase when coming back from PayPal 
    ''' </summary> 
    Private Property Business() As String
        Get
            Return _business
        End Get
        Set(ByVal value As String)
            _business = value
        End Set
    End Property


    ''' <summary> 
    ''' Unique transaction ID generated by PayPal. Helpful to use for checking 
    ''' against fraud to make sure the transaction hasn't already occured. 
    ''' </summary> 
    Private Property TXN_ID() As String
        Get
            Return _txnID
        End Get
        Set(ByVal value As String)
            _txnID = value
        End Set
    End Property

    ''' <summary> 
    ''' Type of transaction from the customer. Possible values are 
    ''' "cart", "express_checkout", "send_money", "virtual_terminal", "web-accept" 
    ''' </summary> 
    Private Property TXN_Type() As String
        Get
            Return _txnType
        End Get
        Set(ByVal value As String)
            _txnType = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the status of the payment from the Customer.Possible values are: 
    ''' "Canceled_Reversal", "Completed", "Denied", "Expired", "Failed", "Pending", 
    ''' "Processed", "Refunded", "Reversed", "Voided" 
    ''' </summary> 
    Private Property PaymentStatus() As String
        Get
            Return _paymentStatus
        End Get
        Set(ByVal value As String)
            _paymentStatus = value
        End Set
    End Property

    ''' <summary> 
    ''' Primary email address of you, the recipient, of the payment. 
    ''' </summary> 
    Private Property ReceiverEmail() As String
        Get
            Return _receiverEmail
        End Get
        Set(ByVal value As String)
            _receiverEmail = value
        End Set
    End Property

    ''' <summary> 
    ''' unique account ID of the payment recipient, which is most likely yourself. 
    ''' </summary> 
    Private Property ReceiverID() As String
        Get
            Return _receiverID
        End Get
        Set(ByVal value As String)
            _receiverID = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the item name passed by yourself or if the customer if you let them enter in an item name 
    ''' </summary> 
    Private Property ItemName() As String
        Get
            Return _itemName
        End Get
        Set(ByVal value As String)
            _itemName = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the item number you set for your own tracking purposes. It is not required by PayPal 
    ''' so if you didn't set it most likely will come back blank. 
    ''' </summary> 
    Private Property ItemNumber() As String
        Get
            Return _itemNumber
        End Get
        Set(ByVal value As String)
            _itemNumber = value
        End Set
    End Property

    ''' <summary> 
    ''' Quantity of the item ordered by the customer 
    ''' </summary> 
    Private Property Quantity() As String
        Get
            Return _quantity
        End Get
        Set(ByVal value As String)
            _quantity = value
        End Set
    End Property

    ''' <summary> 
    ''' Quantity of the items in the shopping cart from the Customer 
    ''' </summary> 
    Private Property QuantityCartItems() As String
        Get
            Return _qtyCartItems
        End Get
        Set(ByVal value As String)
            _qtyCartItems = value
        End Set
    End Property

    ''' <summary> 
    ''' Invoice number passed by yourself, if you didn't pass it to PayPal then this is omitted. 
    ''' </summary> 
    Private Property Invoice() As String
        Get
            Return _invoice
        End Get
        Set(ByVal value As String)
            _invoice = value
        End Set
    End Property

    ''' <summary> 
    ''' Custom value passed by yourself with the item. 
    ''' </summary> 
    Private Property Custom() As String
        Get
            Return _custom
        End Get
        Set(ByVal value As String)
            _custom = value
        End Set
    End Property

    ''' <summary> 
    ''' Memo entered in by the customer on PayPal website note field 
    ''' </summary> 
    Private Property Memo() As String
        Get
            Return _memo
        End Get
        Set(ByVal value As String)
            _memo = value
        End Set
    End Property

    ''' <summary> 
    ''' Amount of tax charged on the payment 
    ''' </summary> 
    Private Property Tax() As String
        Get
            Return _tax
        End Get
        Set(ByVal value As String)
            _tax = value
        End Set
    End Property

    ''' <summary> 
    ''' Full USD amount of customer's payment before the PayPal fee is subtracted 
    ''' </summary> 
    Private Property PaymentGross() As String
        Get
            Return _paymentGross
        End Get
        Set(ByVal value As String)
            _paymentGross = value
        End Set
    End Property

    ''' <summary> 
    ''' Date Time stamp created by PayPal in the following format: 
    ''' HH:MM:SS DD Mmm YY, YYYY PST 
    ''' </summary> 
    Private Property PaymentDate() As String
        Get
            Return _paymentDate
        End Get
        Set(ByVal value As String)
            _paymentDate = value
        End Set
    End Property

    ''' <summary> 
    ''' PayPal's transaction fees associated with purchase. 
    ''' </summary> 
    Private Property PaymentFee() As String
        Get
            Return _paymentFee
        End Get
        Set(ByVal value As String)
            _paymentFee = value
        End Set
    End Property


    ''' <summary> 
    ''' This is the email that the customer used on PayPal or that 
    ''' is registered with PayPal 
    ''' </summary> 
    Private Property PayerEmail() As String
        Get
            Return _payerEmail
        End Get
        Set(ByVal value As String)
            _payerEmail = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's phone number 
    ''' </summary> 
    Private Property PayerPhone() As String
        Get
            Return _payerPhone
        End Get
        Set(ByVal value As String)
            _payerPhone = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's company name if they represent a business 
    ''' </summary> 
    Private Property PayerBusinessName() As String
        Get
            Return _payerBusinessName
        End Get
        Set(ByVal value As String)
            _payerBusinessName = value
        End Set
    End Property

    ''' <summary> 
    ''' This variable is only set if the payment_status=Pending. Possible values are the following: 
    ''' "address", "authorization", "echeck", "intl", "multi-currency", "unilateral", "upgrade", 
    ''' "verify", other" 
    ''' </summary> 
    Private Property PendingReason() As String
        Get
            Return _pendingReason
        End Get
        Set(ByVal value As String)
            _pendingReason = value
        End Set
    End Property

    ''' <summary> 
    ''' This is indicated from what is set in your PayPal profile settings 
    ''' </summary> 
    Private Property ShippingMethod() As String
        Get
            Return _shippingMethod
        End Get
        Set(ByVal value As String)
            _shippingMethod = value
        End Set
    End Property

    ''' <summary> 
    ''' Shipping charges associated with the order. 
    ''' </summary> 
    Private Property Shipping() As String
        Get
            Return _shipping
        End Get
        Set(ByVal value As String)
            _shipping = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's First Name 
    ''' </summary> 
    Private Property PayerFirstName() As String
        Get
            Return _firstName
        End Get
        Set(ByVal value As String)
            _firstName = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's Last Name 
    ''' </summary> 
    Private Property PayerLastName() As String
        Get
            Return _lastName
        End Get
        Set(ByVal value As String)
            _lastName = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's street address 
    ''' </summary> 
    Private Property PayerAddress() As String
        Get
            Return _address
        End Get
        Set(ByVal value As String)
            _address = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's city 
    ''' </summary> 
    Private Property PayerCity() As String
        Get
            Return _city
        End Get
        Set(ByVal value As String)
            _city = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer state of address 
    ''' </summary> 
    Private Property PayerState() As String
        Get
            Return _state
        End Get
        Set(ByVal value As String)
            _state = value
        End Set
    End Property

    ''' <summary> 
    ''' Zip code of customer's address 
    ''' </summary> 
    Private Property PayerZipCode() As String
        Get
            Return _zip
        End Get
        Set(ByVal value As String)
            _zip = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's country 
    ''' </summary> 
    Private Property PayerCountry() As String
        Get
            Return _country
        End Get
        Set(ByVal value As String)
            _country = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's 2 character country code 
    ''' </summary> 
    Private Property PayerCountryCode() As String
        Get
            Return _countryCode
        End Get
        Set(ByVal value As String)
            _countryCode = value
        End Set
    End Property

    ''' <summary> 
    ''' The the address provided is either confirmed or uncomfirmed from PayaPal. Possible values from PayPal 
    ''' are going to be "confirmed" or "unconfirmed" 
    ''' </summary> 
    Private Property PayerAddressStatus() As String
        Get
            Return _addressStatus
        End Get
        Set(ByVal value As String)
            _addressStatus = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer either had a verified or unverified account with PayPal. 
    ''' Possible return values from PayPal are "verified" or "unverified" 
    ''' </summary> 
    Private Property PayerStatus() As String
        Get
            Return _payerStatus
        End Get
        Set(ByVal value As String)
            _payerStatus = value
        End Set
    End Property

    ''' <summary> 
    ''' Customer's unique ID 
    ''' </summary> 
    Private Property PayerID() As String
        Get
            Return _payerID
        End Get
        Set(ByVal value As String)
            _payerID = value
        End Set
    End Property

    ''' <summary> 
    ''' Type of payment from Customer. Possible values from PayPal are "echeck" and "instant" 
    ''' </summary> 
    Private Property PaymentType() As String
        Get
            Return _paymentType
        End Get
        Set(ByVal value As String)
            _paymentType = value
        End Set
    End Property

    ''' <summary> 
    ''' This is the version number of the IPN that makes the post. 
    ''' </summary> 
    Private Property NotifyVersion() As String
        Get
            Return _notifyVersion
        End Get
        Set(ByVal value As String)
            _notifyVersion = value
        End Set
    End Property

    ''' <summary> 
    ''' An encrypted string that is used to validate the transaction. You don't have to use this for anything 
    ''' unless you want to keep it and store it for your records. 
    ''' </summary> 
    Private Property VerifySign() As String
        Get
            Return _verifySign
        End Get
        Set(ByVal value As String)
            _verifySign = value
        End Set
    End Property

#End Region

#Region "Make HTTP POST"

    ''' <summary> 
    ''' This makes the post back to PayPal to verify the order. 
    ''' </summary> 
    Public Sub MakeHttpPost()
        Dim req As HttpWebRequest = DirectCast(WebRequest.Create(Me.PostUrl), HttpWebRequest)

        req.Method = "POST"
        req.ContentLength = Me.RequestLength.Length + 21
        req.ContentType = "application/x-www-form-urlencoded"
        Dim param As Byte() = HttpContext.Current.Request.BinaryRead(HttpContext.Current.Request.ContentLength)
        Me.RequestLength = Encoding.ASCII.GetString(param)
        Me.RequestLength += "&cmd=_notify-validate"
        req.ContentLength = Me.RequestLength.Length

        Dim streamOut As New StreamWriter(req.GetRequestStream(), System.Text.Encoding.ASCII)
        streamOut.Write(Me.RequestLength)
        streamOut.Close()
        Dim streamIn As New StreamReader(req.GetResponse().GetResponseStream())
        Me.Response = streamIn.ReadToEnd()
        streamIn.Close()
    End Sub

#End Region

#Region "Check Status of Order"

    ''' <summary> 
    ''' This checks the status of the order and notifies you via email the status. 
    ''' </summary> 
    Public Sub CheckStatus()
        Select Case Me.Response
            Case Is = "VERIFIED"
                Select Case Me.PaymentStatus
                    Case Is = "Completed"

                        'insert a record into the appropriate tables
                        Dim cn As SqlConnection
                        Dim rs As SqlCommand
                        'Dim rsReply As SqlDataReader
                        cn = New SqlConnection
                        cn.ConnectionString = connStr
                        cn.Open()
                        SQL = "INSERT INTO RATING_REQUEST_TBL(PERSON_ID) " _
                        & "SELECT '-1'"
                        rs = New SqlCommand(SQL, cn)
                        rs.ExecuteNonQuery()
                        Exit Sub

                        If Me.ReceiverEmail = Me.ToEmail Then
                            Select Case Me.TXN_Type
                                Case Is = "cart"
                                    Me.EmailUs("PayPal: Successful Order from Cart")
                                    Exit Select
                                Case Is = "express_checkout"
                                    Me.EmailUs("PayPal: Successful Order from Express Checkout")
                                    Exit Select
                                Case Is = "send_money"
                                    Me.EmailUs("PayPal: Successful Order from Send Money")
                                    Exit Select
                                Case Is = "virtual_terminal"
                                    Me.EmailUs("PayPal: Successful Order from Virtual Terminal")
                                    Exit Select
                                Case Is = "web_accept"
                                    Me.EmailUs("PayPal: Successful Order from Web_Accept")
                                    Exit Select
                                Case Else
                                    Me.EmailUs("PayPal: Order has been placed")
                                    Exit Select
                            End Select
                        Else
                            Me.EmailUs("PayPal: Unknown order...please check your paypal account")
                        End If
                        Exit Select
                    Case Is = "Pending"
                        Exit Sub
                        Select Case Me.PendingReason
                            Case Is = "address"
                                Me.EmailUs("PayPal: Pending Order because of address")
                                Exit Select
                            Case Is = "authorization"
                                Me.EmailUs("PayPal: Pending Order because of authorization")
                                Exit Select
                            Case Is = "echeck"
                                Me.EmailUs("PayPal: Pending Order because of echeck")
                                Exit Select
                            Case Is = "intl"
                                Me.EmailUs("PayPal: Pending Order because of non-US Acccount")
                                Exit Select
                            Case Is = "multi-currency"
                                Me.EmailUs("PayPal: Pending Order because of multi-currency")
                                Exit Select
                            Case Is = "unilateral"
                                Me.EmailUs("PayPal: Pending Order because of Unilateral")
                                Exit Select
                            Case Is = "upgrade"
                                Me.EmailUs("PayPal: Pending Order because of Upgrade")
                                Exit Select
                            Case Is = "verify"
                                Me.EmailUs("PayPal: Pending Order because of Verification needed")
                                Exit Select
                            Case Is = "other"
                                Me.EmailUs("PayPal: Pending Order because of other reason")
                                Exit Select
                            Case Else
                                Me.EmailUs(String.Format("PayPal: Pending Order because of unknown reason of {0}", Me.PendingReason))
                                Exit Select
                        End Select
                        Exit Select
                    Case Is = "Failed"
                        Exit Sub
                        Me.EmailUs("PayPal: Failed order")
                        Exit Select
                    Case Is = "Denied"
                        Exit Sub
                        Me.EmailUs("PayPal: Denied order")
                        Exit Select
                End Select

                Me.EmailBuyer("Order Received", "Your order has been received and will begin processing shortly.")

                Exit Select
            Case Is = "INVALID"
                'insert a record into the appropriate tables
                Dim cn As SqlConnection
                Dim rs As SqlCommand
                'Dim rsReply As SqlDataReader
                cn = New SqlConnection
                cn.ConnectionString = connStr
                cn.Open()
                SQL = "INSERT INTO RATING_REQUEST_TBL(PERSON_ID) " _
                & "SELECT '-2'"
                rs = New SqlCommand(SQL, cn)
                rs.ExecuteNonQuery()
                Exit Sub

                Me.EmailUs("PayPal: Invalid order, please review and investigate")
                Exit Select
            Case Else
                'insert a record into the appropriate tables
                Dim cn As SqlConnection
                Dim rs As SqlCommand
                'Dim rsReply As SqlDataReader
                cn = New SqlConnection
                cn.ConnectionString = connStr
                cn.Open()
                SQL = "INSERT INTO RATING_REQUEST_TBL(PERSON_ID) " _
                & "SELECT '-3'"
                rs = New SqlCommand(SQL, cn)
                rs.ExecuteNonQuery()
                Exit Sub
                Me.EmailUs("PayPal: ERROR, response is " + Me.Response)
                Exit Select
        End Select
    End Sub


#End Region

#Region "Mail Company the Order"
    ''' <summary> 
    ''' Email yourself/company the order. This requires a subject line. Make sure to set SMTP properties of the PayPal object 
    ''' and the FromEmail and ToEmail properties as well. 
    ''' </summary> 
    ''' <param name="subject"></param> 
    Private Sub EmailUs(ByVal subject As String)
        Dim mailObj As New MailMessage()
        mailObj.From = New MailAddress(Me.FromEmail)
        mailObj.Subject = subject
        mailObj.[To].Add(New MailAddress(Me.ToEmail))
        mailObj.IsBodyHtml = True
        mailObj.Body = "<br />" + "Transaction ID: " + Me.TXN_ID + "<br />" + "Transaction Type:" + Me.TXN_Type + "<br />" + "Payment Type: " + Me.PaymentType + "<br />" + "Payment Status: " + Me.PaymentStatus + "<br />" + "Pending Reason: " + Me.PendingReason + "<br />" + "Payment Date: " + Me.PaymentDate + "<br />" + "Receiver Email: " + Me.ReceiverEmail + "<br />" + "Invoice: " + Me.Invoice + "<br />" + "Item Number: " + Me.ItemNumber + "<br />" + "Item Name: " + Me.ItemName + "<br />" + "Quantity: " + Me.Quantity + "<br />" + "Custom: " + Me.[Custom] + "<br />" + "Payment Gross: " + Me.PaymentGross + "<br />" + "Payment Fee: " + Me.PaymentFee + "<br />" + "Payer Email: " + Me.PayerEmail + "<br />" + "First Name: " + Me.PayerFirstName + "<br />" + "Last Name: " + Me.PayerLastName + "<br />" + "Street Address: " + Me.PayerAddress + "<br />" + "City: " + Me.PayerCity + "<br />" + "State: " + Me.PayerState + "<br />" + "Zip Code: " + Me.PayerZipCode + "<br />" + "Country: " + Me.PayerCountry + "<br />" + "Address Status: " + Me.PayerAddressStatus + "<br />" + "Payer Status: " + Me.PayerStatus + "<br />" + "Verify Sign: " + Me.VerifySign + "<br />" + "Notify Version: " + Me.NotifyVersion + "<br />"

        Dim objSmtp As New SmtpClient()

        objSmtp.Host = Me.SmtpHost
        objSmtp.Port = System.Int32.Parse(Me.SmtpPort)
        objSmtp.UseDefaultCredentials = False
        objSmtp.Credentials = New System.Net.NetworkCredential(Me.FromEmail, Me.FromEmailPassword)
        objSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
        objSmtp.Send(mailObj)
    End Sub

#End Region

#Region "Mail the Customer the Order details"

    Private Sub EmailBuyer(ByVal subject As String, ByVal message As String)
        Dim mailObj As New MailMessage()
        mailObj.From = New MailAddress(Me.FromEmail)
        mailObj.Subject = subject
        mailObj.Body = message
        mailObj.[To].Add(New MailAddress(Me.PayerEmail))
        mailObj.IsBodyHtml = True

        Dim objSmtp As New SmtpClient()

        objSmtp.Host = Me.SmtpHost
        objSmtp.Port = System.Int32.Parse(Me.SmtpPort)
        objSmtp.UseDefaultCredentials = False
        objSmtp.Credentials = New System.Net.NetworkCredential(Me.FromEmail, Me.FromEmailPassword)
        objSmtp.DeliveryMethod = SmtpDeliveryMethod.Network
        objSmtp.Send(mailObj)
    End Sub

#End Region

#Region "Fill Properties"

    Private Sub fillProperties()
        Me.RequestLength = HttpContext.Current.Request.Form.ToString()

        Me.PayerCity = HttpContext.Current.Request.Form("address_city")
        Me.PayerCountry = HttpContext.Current.Request.Form("address_country")
        Me.PayerCountryCode = HttpContext.Current.Request.Form("address_country_code")
        Me.PayerState = HttpContext.Current.Request.Form("address_state")
        Me.PayerAddressStatus = HttpContext.Current.Request.Form("address_status")
        Me.PayerAddress = HttpContext.Current.Request.Form("address_street")
        Me.PayerZipCode = HttpContext.Current.Request.Form("address_zip")
        Me.PayerFirstName = HttpContext.Current.Request.Form("first_name")
        Me.PayerLastName = HttpContext.Current.Request.Form("last_name")
        Me.PayerBusinessName = HttpContext.Current.Request.Form("payer_business_name")
        Me.PayerEmail = HttpContext.Current.Request.Form("payer_email")
        Me.PayerID = HttpContext.Current.Request.Form("payer_id")
        Me.PayerStatus = HttpContext.Current.Request.Form("payer_status")
        Me.PayerPhone = HttpContext.Current.Request.Form("contact_phone")
        Me.Business = HttpContext.Current.Request.Form("business")
        Me.ItemName = HttpContext.Current.Request.Form("item_name")
        Me.ItemNumber = HttpContext.Current.Request.Form("item_number")
        Me.Quantity = HttpContext.Current.Request.Form("quantity")
        Me.ReceiverEmail = HttpContext.Current.Request.Form("receiver_email")
        Me.ReceiverID = HttpContext.Current.Request.Form("receiver_id")
        Me.[Custom] = HttpContext.Current.Request.Form("custom")
        Me.Memo = HttpContext.Current.Request.Form("memo")
        Me.Invoice = HttpContext.Current.Request.Form("invoice")
        Me.Tax = HttpContext.Current.Request.Form("tax")
        Me.QuantityCartItems = HttpContext.Current.Request.Form("num_cart_items")
        Me.PaymentDate = HttpContext.Current.Request.Form("payment_date")
        Me.PaymentStatus = HttpContext.Current.Request.Form("payment_status")
        Me.PaymentType = HttpContext.Current.Request.Form("payment_type")
        Me.PendingReason = HttpContext.Current.Request.Form("pending_reason")
        Me.TXN_ID = HttpContext.Current.Request.Form("txn_id")
        Me.TXN_Type = HttpContext.Current.Request.Form("txn_type")
        Me.PaymentFee = HttpContext.Current.Request.Form("mc_fee")
        Me.PaymentGross = HttpContext.Current.Request.Form("mc_gross")
        Me.NotifyVersion = HttpContext.Current.Request.Form("notify_version")
        Me.VerifySign = HttpContext.Current.Request.Form("verify_sign")
    End Sub

#End Region

End Class
