<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,nom,email,coments


nom=Request.Form("usuario")
email=Request.Form("correo")
coments=Request.Form("coments")

set conn=Server.CreateObject("ADODB.Connection")
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\dgoou\OneDrive\Desktop\pagina web\Nueva carpeta\COMPRAS.mdb	"
conn.execute "INSERT INTO comentarios(usuario,correo,coments) values('" & nom & "','" & email &  "','" & coments & "')"
conn.close
set conn=nothing
response.redirect("COMENTARIOS.html")
%>
</body>
</html>