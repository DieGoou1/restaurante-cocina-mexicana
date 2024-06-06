<html>
<body>
<%@ LANGUAGE="VBSCRIPT" %>
<%
option explicit
Dim conn,limon,bayleys,naranja,jamaica,horchata,azulito,mojito,tacos,guacamole,mole,enchiladas,chiles_rellenos,pozole,tamales,chilaquiles,tostadas,sopes,birria,mql1,mql2,mql3,mql4,mql5
Dim id,cantidad,codigop,precio


id=Request.Form("id")
limon=Request.Form("limon")
bayleys=Request.Form("bayleys")
naranja=Request.Form("naranja")
jamaica=Request.form("jamaica")
horchata=Request.Form("horchata")
azulito=Request.Form("azulito")
tacos=Request.form("tacos") 
guacamole=Request.Form("guacamole")
mole=Request.Form("mole")
enchiladas=Request.Form("enchiladas")
chiles_rellenos=Request.Form("chiles_rellenos")
pozole=Request.Form("pozole")
tamales=Request.Form("tamales")
chilaquiles=Request.Form("chilaquiles")
tostadas=Request.Form("tostadas")
sopes=Request.Form("sopes")
birria=Request.Form("birria")
mql1=Request.Form("mql1")
mql2=Request.Form("mql2")
mql3=Request.Form("mql3")
mql4=Request.Form("mql4")
mql5=Request.Form("mql5")
id=Request.Form("id")
cantidad=Request.Form("cantidad")
codigop=Request.Form("codigop")
precio=Request.Form("precio")


Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\dgoou\OneDrive\Desktop\pagina web\Nueva carpeta\COMPRAS.mdb"


if limon= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'limon',30)"
end if
if bayleys= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'bayleys',80)"
end if
if naranja= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'naranja',30)"
end if
if jamaica= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'jamaica',30)"
end if
if horchata= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'horchata',30)"
end if
if azulito= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'azulito',60)"
end if
if mojito= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mojito',80)"
end if
if tacos= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'tacos',10)"
end if
if guacamole= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'guacamole',8)"
end if
if mole= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mole',12)"
end if
if enchiladas= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'enchiladas',9)"
end if
if chiles_rellenos= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'chiles_rellenos',11)"
end if
if pozole= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'pozole',30)"
end if
if tamales= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'tamales',35)"
end if
if chilaquiles= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'chilaquiles',30)"
end if
if mojito= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mojito',40)"
end if
if tostadas= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'tostadas',40)"
end if
if sopes= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'sopes',60)"
end if
if birria= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'birria',70)"
end if
if mql1= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mql1',220)"
end if
if mql2= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mql2',180)"
end if
if mql3= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mql3',522)"
end if
if mql4= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mql4',350)"
end if
if mql5= 1 then
conn.execute "INSERT INTO PEDIDOS(id,cantidad,codigop,precio) values(1,"& CInt(cantidad) &",'mql5',160)"
end if
conn.close
set conn=nothing
response.redirect("pagina1.html")
%>
</body>
</html>