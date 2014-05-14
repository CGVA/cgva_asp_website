<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DraftPictures.aspx.vb" Inherits="DraftPictures" enableEventValidation="false" %>
<% System.Web.HttpContext.Current.Response.AddHeader( "Cache-Control","no-cache")
System.Web.HttpContext.Current.Response.Expires = 0
    System.Web.HttpContext.Current.Response.Cache.SetNoStore()
System.Web.HttpContext.Current.Response.AddHeader( "Pragma", "no-cache")%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:Panel ID="picturePanel" runat="server"></asp:Panel>
    </div>
    </form>
</body>
</html>
