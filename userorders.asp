<% @language = "vbscript" %>
<% Option Explicit %>
<% Response.Buffer = "false" %>
<html>
    <body>
        <%
        dim conn,res,usn
        set conn=Server.CreateObject("ADODB.Connection")
        conn.provider="Microsoft.Jet.OLEDB.4.0"
        conn.open "C:/inetpub/wwwroot/WT_MiniProject/Foodify/DB.mdb"
        set res = Server.CreateObject("ADODB.Recordset")
        res.Open "OrderTable",conn,,,2
        Do while not res.EOF
                        if(u = usn And p = pass) Then
                            status = True
                                'Diplay user that the login was sucessfull & redirect to the home page
                                Response.write("<script language=""javascript"">alert('Order has been placed for');</script>")
                                Server.Execute("menu.html")  
                            Exit Do
                        else
                            res.MoveNext
                            status = False
                        End if
                    Loop
        
        
        %>
    </body>
</html>