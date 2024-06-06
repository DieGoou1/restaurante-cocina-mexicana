<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn, nom, primer_apellido, segundo_apellido,correo, mensaje

NOM=Request.Form("nombre")
PRIMER_APELLIDO=Request.Form("primer_apellido")
SEGUNDO_APELLIDO=Request.Form("segundo_apellido")
correo=Request.Form("correo")
mensaje=Request.Form("mensaje")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\dgoou\OneDrive\Desktop\pagina web\Nueva carpeta\COMPRAS.mdb"
conn.Execute "INSERT INTO COMPRAS(nombre,primer_apellido,segundo_apellido,correo,mensaje) VALUES ('" & nom & "','" & primer_apellido & "','" & segundo_apellido & "','" & correo & "','" & mensaje & "')"
conn.close
set conn=nothing
response.redirect("FORMULARIO.html")
%>
</body>
</html>