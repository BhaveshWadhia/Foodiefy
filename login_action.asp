<% @language="vbscript" %>
<% Option Explicit %>
<% Response.Buffer = "false" %>

<%  
dim conn,res,u,p,usn,pass,status,result
set conn=Server.CreateObject("ADODB.Connection")
conn.provider="Microsoft.Jet.OLEDB.4.0"
conn.open "C:/inetpub/wwwroot/WT_MiniProject/Foodify/DB.mdb"
u = Request.Form("uname")
p = Request.Form("psw")
if u =" " And p=" " Then
    Response.Write("Fileds are empty")
else
    set res = Server.CreateObject("ADODB.Recordset")
    res.Open "LoginTable",conn,,,2
      Do while not res.EOF
            usn = res("UserName")
            pass = res("Password")
            if(u = usn And p = pass) Then
                MsgBox "Login Successfull"
                status = True
                Exit Do
            else
                res.MoveNext
                status = False
            End if
        Loop
    if res.EOF = True And status = False Then
        MsgBox "Login Failed"
    End If
End If
%>

