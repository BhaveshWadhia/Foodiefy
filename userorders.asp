<% @language = "vbscript" %>
<% Option Explicit %>
<% Response.Buffer = "false" %>
<html>
    <body>
        <%
            dim conn,usn,odr,res
            Set conn= Server.CreateObject("ADODB.Connection")
            conn.Provider = "Microsoft.Jet.OLEDB.4.0"
            conn.Open "C:\inetpub\wwwroot\WT_MiniProject\Foodify\DB.mdb"
            Set res = Server.CreateObject("ADODB.RecordSet")
            res.open "OrderTable", conn, 0, 3, 2
            usn = Request.Form("username")
            odr = Request.Form("orders")
            'Place the users order into the database
            res.AddNew()
            res("UserName") = usn
            res("OrderDetails") = odr
            res.Update()
            res.MoveNext
            conn.close
        Response.write("<script language=""javascript"">alert('Order has been placed');</script>")
        Server.Execute("menu.html")  
        set conn = nothing
        %>
    </body>
</html>