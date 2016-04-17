
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>随机日记</title>
</head>

<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../信管班纪念册.mdb")
conn.Open strConn 

dim rs,sql,ID
set rs=server.CreateObject("adodb.recordset")
sql="select * from 随机日记表"
rs.open sql,conn,1,2
do while not rs.eof
response.write"</br><table width='290px'><tr><td width='30%'>日记人：</td><td width='70%' align='left'>"&rs("姓名")&"</td></tr><tr><td width='30%'>日记内容：</td><td width='70%'align='left' >"&rs("日记内容")&"</td></tr></table>"
rs.movenext
loop
rs.close
set rs=nothing

%>
</body>
</html>
