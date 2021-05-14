<% @language="vbscript" %>
<% option explicit %>
<html>
    <body>
        <%
            Dim conn,res
            Set conn= Server.CreateObject("ADODB.Connection")
            conn.Provider = "Microsoft.Jet.OLEDB.4.0"
            conn.Open "C:\inetpub\wwwroot\WT_MiniProject\Foodify\DB.mdb"
            Set res = Server.CreateObject("ADODB.RecordSet")
            res.open "Contact", conn, 0, 3, 2
            res.AddNew()
            res("Myname") = Request.form("uname") 
            res("Email") = Request.form("email")
            res("Subject") = Request.form("subject")
            res("Message") = Request.form("message")

            res.Update()  
            res.MoveNext
            conn.close

            Response.write("<script language=""javascript"">alert('We have successfully received your message');</script>")
            Server.Execute("contact.html")  
            set conn = nothing
        %>
    </body>
</html>




