<!-- #include virtual = "/incs/dbConnection.inc" -->
<%
            if request("registrationType") = "TEAM" then
                teamName = Replace(request("teamName"),"'","''")
                division = request("division")
                teamFee = request("teamFee")
                additionalPlayers = request("additionalPlayers")
                totalFee = CInt(teamFee) + CInt(additionalPlayers)
                
                'VERIFY TEAM NAME DOESN'T ALREADY EXIST FOR THIS DIVISION
                SQL = "SELECT COUNT(*) FROM LASTDIG_TEAM_TBL "_
                        & "WHERE TEAM_NAME='" & teamName & "' "_
                        & "AND DIVISION = '" & division & "'"
                set rs = cn.Execute(SQL)
                if rs(0) > 0 then
                    Session("Err") = "The team name you entered already exists in the selected division. Please enter a new name to continue with registration."
                    cn.Close
                    Response.Redirect("paypal.asp")
                end if
                
                SQL = "Set Nocount on "_
                    & "INSERT INTO LASTDIG_TEAM_TBL(TEAM_NAME,DIVISION,FEE_PAID) "_
                    & "SELECT '" & teamName & "', "_
                    & "'" & division & "', "_
                    & "'" & totalFee & "' "_
                    & "SELECT @@IDENTITY "_
                    & "Set Nocount OFF"
                '& "'" & totalFee & "'"
                
                'Response.Write(SQL)
                set rs = cn.Execute(SQL)
                newID = rs(0)
                'Response.Write(newID)
                'Response.End
            elseif request("registrationType") = "PLAYER" then
                teamNameTemp = Replace(request("teamName"),"'","''")
                 teamNameArray = Split(teamNameTemp," - ")
                teamName = teamNameArray(0)
                division = teamNameArray(1)
                additionalPlayers = request("additionalPlayers")
                totalFee = CInt(additionalPlayers)
                SQL = "Set Nocount on "_
                    & "INSERT INTO LASTDIG_TEAM_TBL(TEAM_NAME,DIVISION,FEE_PAID) "_
                    & "SELECT '" & teamName & "', "_
                    & "'" & division & "', "_
                    & "'" & totalFee & "' "_
                    & "SELECT @@IDENTITY "_
                    & "Set Nocount OFF"
                
                'Response.Write(SQL)
                set rs = cn.Execute(SQL)
                newID = rs(0)
                Response.Write(newID)
                Response.End
            else
                'ERROR
            end if

Response.Redirect("PayPal.aspx")
Response.End
            cmd = "_xclick"
            business = "vp@cgva.org"
            item_name = "Last Dig In Denver 2010"         
            item_number = newID
            amount = totalFee
            no_shipping = "1"

            return_url = "http://lastdigindenver.org/paypal.asp"
            'If AppSettings("SendToReturnURL").ToString = "true" Then
            '    rm = "2"
            'Else
                rm = "1"
            'End If
           
            notify_url = "http://cgva.org/MyCGVA/IPNHandler.aspx"
            cancel_url = "http://lastdigindenver.org/paypal.asp"
            currency_code = "USD"
       
           'URL = "https://www.sandbox.paypal.com/cgi-bin/webscr"
           'paypaluser = "jcross_1250184916_biz_api1.aol.com"
            'pwd = "1250184926"
            'signature = "AUcDthwlq0Is0GjmLHeaEz1kcu.YABJyNUCXUNNwhS9mxYFw5vrHw.u8"
           
           URL = "https://www.paypal.com/cgi-bin/webscr"
           paypaluser = "vp_api1.cgva.org"
            pwd = "AZU55PW27DFSAPB3"
            signature = "AFcWxV21C7fd0v3bYYYRCpSSRl31AlhtnK28YVt-4kJIY8C3c-r2PpkC"




%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Untitled Page</title>
</head>
<body>
    <form id="payForm" method="post" action="https://www.paypal.com/cgi-bin/webscr">
    
        <input type="hidden" name="cmd" value="<%Response.Write (cmd)%>">
        <input type="hidden" name="business" value="<%Response.Write (business)%>">
        <input type="hidden" name="item_name" value="<%Response.Write (item_name)%>">
        <input type="hidden" name="item_number" value="<%Response.Write (item_number)%>">
        <input type="hidden" name="amount" value="<%Response.Write (amount)%>">
        <input type="hidden" name="no_shipping" value="<%Response.Write (no_shipping)%>">
        <input type="hidden" name="return" value="<%Response.Write (return_url)%>">
        <input type="hidden" name="rm" value="<%Response.Write (rm)%>">
        <input type="hidden" name="notify_url" value="<%Response.Write (notify_url)%>">
        <input type="hidden" name="cancel_return" value="<%Response.Write (cancel_url)%>">
        <input type="hidden" name="currency_code" value="<%Response.Write (currency_code)%>">
        <input type="hidden" name="custom" value="<%Response.Write (newID)%>">
        <input type="hidden" name="user" value="<%Response.Write (paypaluser)%>">
        <input type="hidden" name="pwd" value="<%Response.Write (pwd)%>">
        <input type="hidden" name="signature" value="<%Response.Write (signature)%>">
        
    </form>
    <script type="text/javascript" language="javascript">
    <!--
        document.forms["payForm"].submit();
    -->
    </script>
    
</body>
</html>




