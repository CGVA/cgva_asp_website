<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SampleNewRegistrationEmail.aspx.vb" Inherits="MyCGVA_SampleNewRegistrationEmail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Sample registration email</title>
</head>
<body>
    <form id="form1" runat="server">
    &lt;CGVA LOGO HERE&gt;
    <p />
    <div>
    Welcome &lt;USER NAME HERE&gt;! 
    <p />
    Your MyCGVA user name is &lt;first initial, last name (add a number to this if the name already exists)&gt;.
    <br />
    Your temporary password is &lt;randomly generated password&gt;. When you log in for the first time,
    you will be asked to reset your password before proceeding. If you have any questions, please contact us &lt;link to contact us email&gt;
        or <a href='Login.aspx'>log in</a> now!</div>
    </form>
</body>
</html>
