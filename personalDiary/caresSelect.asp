
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>����ռ�</title>
</head>

<body>
<%
Dim conn,strConn 
Set conn=Server.CreateObject("ADODB.Connection")
strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../�Źܰ�����.mdb")
conn.Open strConn 

dim rs,sql,ID
set rs=server.CreateObject("adodb.recordset")
sql="select * from ����ռǱ�"
rs.open sql,conn,1,2
do while not rs.eof
response.write"</br><table width='290px'><tr><td width='30%'>�ռ��ˣ�</td><td width='70%' align='left'>"&rs("����")&"</td></tr><tr><td width='30%'>�ռ����ݣ�</td><td width='70%'align='left' >"&rs("�ռ�����")&"</td></tr></table>"
rs.movenext
loop
rs.close
set rs=nothing

%>
</body>
</html>
